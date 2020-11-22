using Cassia;
using System;
using System.IO;
using System.Security.Cryptography;

namespace SuncatCommon
{
    public static class SuncatUtilities
    {
        public static string GetMD5HashFromFile(string filename)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash);
                }
            }
        }

        public static ITerminalServicesSession GetActiveSession()
        {
            ITerminalServicesSession activeSession = null;
            var manager = new TerminalServicesManager();

            if (manager != null)
            {
                using (var server = manager.GetLocalServer())
                {
                    server.Open();

                    foreach (var session in server.GetSessions())
                    {
                        if (session.ConnectionState == ConnectionState.Active)
                        {
                            activeSession = session;
                            break;
                        }
                    }
                }
            }

            return activeSession;
        }
    }
}
