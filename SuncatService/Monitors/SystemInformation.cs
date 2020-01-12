using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace SuncatService.Monitors
{
    public static class SystemInformation
    {
        // Get the network interface with the lowest metric because that's the one Windows uses when Automatic Metric feature is enabled.
        private static ManagementObject GetNetworkInterfaceWithLowestMetric()
        {
            ManagementObject value = null;

            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"))
                {
                    var objects = searcher.Get().Cast<ManagementObject>();
                    value = objects.OrderBy(o => o["IPConnectionMetric"]).FirstOrDefault();
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

            return value;
        }

        public static string GetIPv4Address()
        {
            string value = null;

            try
            {
                var values = (string[])GetNetworkInterfaceWithLowestMetric()["IPAddress"];
                value = values.Where(o => o.Contains(".")).FirstOrDefault();
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

            return value;
        }
        
        public static string GetMACAddress()
        {
            string value = null;

            try
            {
                value = Convert.ToString(GetNetworkInterfaceWithLowestMetric()["MACAddress"]);
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

            return value;
        }
    }
}
