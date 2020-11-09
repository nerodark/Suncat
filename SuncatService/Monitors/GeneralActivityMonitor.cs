using SuncatObjects;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SuncatService.Monitors
{
    public static class GeneralActivityMonitor
    {
        private static string publicIP;
        private static Thread heartBeatThread;
        private static Thread getPublicIPThread;

        public static event TrackEventHandler Track;

        private static void HeartBeat()
        {
            while (true)
            {
                var te = new TrackEventArgs();
                var log = new SuncatLog();

                log.DateTime = DateTime.Now;
                log.Event = SuncatLogEvent.HeartBeat;
                te.LogData = log;

                Track?.Invoke(null, te);

                Thread.Sleep(1000 * 60 * 5); // 5 minutes
            }
        }

        private static void GetPublicIP()
        {
            while (true)
            {
                while (true)
                {
                    try
                    {
                        var webRequest = new SuncatWebClient();
                        publicIP = webRequest.DownloadString("https://checkip.amazonaws.com").Trim();

                        var te = new TrackEventArgs();
                        var log = new SuncatLog();

                        log.DateTime = DateTime.Now;
                        log.Event = SuncatLogEvent.GetPublicIP;
                        log.Data1 = publicIP;
                        te.LogData = log;

                        Track?.Invoke(null, te);
                        break;
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
                        
                        Thread.Sleep(1000 * 60 * 5); // repeat every 5 minutes until public IP has been successfully retrieved
                    }
                }

                Thread.Sleep(1000 * 60 * 30); // check every 30 minutes if public IP has changed
            }
        }

        public static void Start()
        {
            heartBeatThread = new Thread(() => HeartBeat());
            heartBeatThread.IsBackground = true;
            heartBeatThread.Start();

            getPublicIPThread = new Thread(() => GetPublicIP());
            getPublicIPThread.IsBackground = true;
            getPublicIPThread.Start();
        }

        public static void Stop()
        {
        }
    }

    public class SuncatWebClient : WebClient
    {
        protected override WebRequest GetWebRequest(Uri address)
        {
            var webRequest = base.GetWebRequest(address);
            webRequest.Timeout = 1000 * 10; // allow 10 seconds to retrieve public IP
            return webRequest;
        }
    }
}
