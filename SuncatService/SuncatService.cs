using Cassia;
using SuncatCommon;
using SuncatService.Monitors;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.ServiceProcess;
using System.Threading;

namespace SuncatService
{
    public partial class SuncatService : ServiceBase
    {
        private BlockingCollection<SuncatLog> logs = new BlockingCollection<SuncatLog>();
        private Thread logProcessor;

        public SuncatService()
        {
            InitializeComponent();
        }

        internal void OnDebug(string[] args)
        {
            OnStart(args);
            Console.ReadLine();
            OnStop();
        }

        protected override void OnStart(string[] args)
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);

            if (!principal.IsInRole(WindowsBuiltInRole.Administrator))
            {
                throw new UnauthorizedAccessException();
            }

            logProcessor = new Thread(() => ProcessLogs());
            logProcessor.IsBackground = true;
            logProcessor.Start();

            GeneralActivityMonitor.Track += Log;
            GeneralActivityMonitor.Start();

            FileSystemActivityMonitor.Track += Log;
            FileSystemActivityMonitor.Start();

            WebActivityMonitor.Track += Log;
            WebActivityMonitor.Start();

            HookActivityMonitor.Track += Log;
            HookActivityMonitor.Start();
        }

        protected override void OnStop()
        {
            GeneralActivityMonitor.Track -= Log;
            GeneralActivityMonitor.Stop();

            FileSystemActivityMonitor.Track -= Log;
            FileSystemActivityMonitor.Stop();

            WebActivityMonitor.Track -= Log;
            WebActivityMonitor.Stop();

            HookActivityMonitor.Track -= Log;
            HookActivityMonitor.Stop();
        }

        private void ProcessLogs()
        {
            SuncatLog log;

            while (logs.TryTake(out log, -1))
            {
                var manager = new TerminalServicesManager();

                if (manager != null && manager.ActiveConsoleSession != null && !string.IsNullOrEmpty(manager.ActiveConsoleSession.UserName))
                {
                    var text =
                        $"[{log.DateTime}] {log.Event}: " +
                        (log.Data1 != null ? log.Data1.Trim() : string.Empty) +
                        (log.Data2 != null ? (log.Data1 != null ? " -> " : string.Empty) + log.Data2.Trim() : string.Empty) +
                        (log.Data3 != null ? (log.Data1 != null || log.Data2 != null ? " -> " : string.Empty) + log.Data3.Trim() : string.Empty) +
                        Environment.NewLine;

                    // To view in DebugView (SysInternals)
                    //Download DebugView
                    //DebugView for Windows is available for download from Microsoft’s Sysinternals team.

                    //Emitting debug information from your code
                    //To make your code tell DebugView what it’s doing you use the Debug class in the System.Diagnostics namespace like so:

                    //Debug.WriteLine("[TS] Generating language object for script file");
                    //This will make the following appear in DebugView as the code executes:

                    //image

                    //Note that the Visual Studio debugger can not be attached, otherwise it’ll grab all debug output to the Output window.

                    //Also, note that the Debug class has multiple methods with multiple overloads to simplify string output.

                    //Start listening to debug output
                    //Start DebugView and make sure to run it as an administrator:

                    //image

                    //Start DebugView and make sure Capture Events is enabled.It’s available on the Capture menu, or through the tool button:
                    //image

                    //Next, make sure DebugView is set to Capture Global Win32:

                    //image

                    //Side note: I’m not entirely sure why this is required, but my guess is that dbgview.exe is run in the context of a user (Session ID: 0) and w3wp.exe is run in the context of a service (Session ID: 1).

                    //Filtering and highlighting
                    //Depending on what other processes are emitting debug info, you might want to apply filters and/or highlighting to the list to avoid being overwhelmed. :)

                    //Simply click Filter/Highlight on the Edit menu to apply filters and highlighting.

                    //image

                    //The filter above would ensure we only include debug messages starting with [TS]:

                    //image

                    //Keyboard shortcuts
                    //CTRL + X (clear the list)
                    //CTRL + T(switch time stamp format)
                    //CTRL + E(enable/disable capture)
                    //Noteworthy
                    //Here’s the setup on which this post is based:

                    //Windows 7 64-bit
                    //IIS7 or Cassini
                    //IIS7 application pool run as NETWORK SERVICE or a network domain account
                    //Listening to debug output from a remote computer
                    //You can essentially listen to any computer you have TCP/IP access to (note that DebugView has to be running on the target machine). Click Connect on the Computer menu and enter the host name or IP address of the computer you want to listen to. If you want to run DebugView in a minimized, unobtrusive way you can use several different arguments when starting it. Start a command prompt and run dbgview.exe /? to get a list of all available options.
                    #if DEBUG
                        Debug.WriteLine(text); 
                    #else
                        Trace.WriteLine(text);
                    #endif

                    //InsertToDatabase(manager, log);

                    if (Environment.UserInteractive)
                    {
                        Console.WriteLine(text);
                    }
                    else
                    {
                        try
                        {
                            File.AppendAllText($@"C:\Insights.txt", text);
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
                }
            }
        }

        public void Log(object sender, TrackEventArgs e)
        {
            // Ignore logs created by this application.
            if (   (e.LogData.Data1 == null || (e.LogData.Data1 != null && !e.LogData.Data1.Contains("Insights.txt")))
                && (e.LogData.Data2 == null || (e.LogData.Data2 != null && !e.LogData.Data2.Contains("Insights.txt")))
                && (e.LogData.Data3 == null || (e.LogData.Data3 != null && !e.LogData.Data3.Contains("Insights.txt"))))
            {
                logs.Add(e.LogData);
            }
        }

        private async void InsertToDatabase(TerminalServicesManager manager, SuncatLog log)
        {
            try
            {
                using (var connection = new SqlConnection("Data Source=server;Database=master;User ID=user;Password=password"))
                {
                    await connection.OpenAsync();

                    //if (log.Event == SuncatLogEvent.OpenURL)
                    //{
                    //    connection.ChangeDatabase("ip2location");

                    //    // speed up the query by forcing SQL Server to use the indexes
                    //    using (var command = new SqlCommand(@"
                    //        SELECT
                    //            country_code,
                    //            country_name,

                    //            region_name,
                    //            city_name,
                    //            latitude,
                    //            longitude,
                    //            zip_code,
                    //            time_zone
                    //        FROM ip2location_db11
                    //        WHERE ip_from = (
                    //            SELECT MAX(ip_from) 
                    //            FROM ip2Location_db11
                    //            WHERE ip_from <= @Ip
                    //        ) 
                    //        AND ip_to = ( 
                    //            SELECT MIN(ip_to) 
                    //            FROM ip2Location_db11
                    //            WHERE ip_to >= @Ip
                    //        )
                    //    ", connection))
                    //    {
                    //        try
                    //        {
                    //            var host = new Uri(log.Data1).DnsSafeHost;
                    //            var ip = Dns.GetHostAddresses(host).First();
                    //            remoteIp = ip.ToString();
                    //            var bytes = ip.GetAddressBytes();
                    //            var ip2int = (uint)IPAddress.NetworkToHostOrder((int)BitConverter.ToUInt32(bytes, 0));

                    //            var ipParam = new SqlParameter("@Ip", SqlDbType.BigInt);
                    //            ipParam.Value = ip2int;
                    //            command.Parameters.Add(ipParam);

                    //            using (var dataReader = await command.ExecuteReaderAsync())
                    //            {
                    //                if (await dataReader.ReadAsync())
                    //                {
                    //                    countryCode = dataReader.GetString(0);
                    //                    countryName = dataReader.GetString(1);

                    //                    regionName = dataReader.GetString(2);
                    //                    cityName = dataReader.GetString(3);
                    //                    latitude = dataReader.GetDouble(4);
                    //                    longitude = dataReader.GetDouble(5);
                    //                    zipCode = dataReader.GetString(6);
                    //                    timeZone = dataReader.GetString(7);
                    //                }
                    //            }
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            #if DEBUG
                    //                Debug.WriteLine(ex);

                    //                if (ex.InnerException != null)
                    //                    Debug.WriteLine(ex.InnerException);
                    //            #else
                    //                Trace.WriteLine(ex);

                    //                if (ex.InnerException != null)
                    //                    Trace.WriteLine(ex.InnerException);
                    //            #endif
                    //        }
                    //    }
                    //}

                    connection.ChangeDatabase("Suncat");

                    using (var command = new SqlCommand(@"
                        INSERT INTO SuncatLogs (
                            DATETIME,
                            EVENT,
                            DATA1,
                            DATA2,
                            DATA3,
                            COMPUTERNAME,
                            USERNAME,
                            LOCALIP,
                            MACADDRESS
                        ) VALUES (
                            @DateTime,
                            @Event,
                            @Data1,
                            @Data2,
                            @Data3,
                            @ComputerName,
                            @UserName,
                            @LocalIp,
                            @MacAddress
                        )
                    ", connection))
                    {
                        var data1 = log.Data1;
                        var data2 = log.Data2;
                        var data3 = log.Data3;
                        var localIp = SystemInformation.GetIPv4Address();
                        var macAddress = SystemInformation.GetMACAddress();

                        command.Parameters.AddWithValue("@DateTime", log.DateTime);
                        command.Parameters.AddWithValue("@Event", log.Event.ToString());
                        command.Parameters.AddWithValue("@Data1", string.IsNullOrWhiteSpace(data1) ? (object)DBNull.Value : data1);
                        command.Parameters.AddWithValue("@Data2", string.IsNullOrWhiteSpace(data2) ? (object)DBNull.Value : data2);
                        command.Parameters.AddWithValue("@Data3", string.IsNullOrWhiteSpace(data3) ? (object)DBNull.Value : data3);
                        command.Parameters.AddWithValue("@ComputerName", Environment.MachineName ?? string.Empty);
                        command.Parameters.AddWithValue("@UserName", manager.ActiveConsoleSession.UserName ?? string.Empty);
                        command.Parameters.AddWithValue("@LocalIp", string.IsNullOrEmpty(localIp) ? (object)DBNull.Value : localIp);
                        command.Parameters.AddWithValue("@MacAddress", string.IsNullOrEmpty(macAddress) ? (object)DBNull.Value : macAddress);

                        await command.ExecuteNonQueryAsync();
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
    }
}
