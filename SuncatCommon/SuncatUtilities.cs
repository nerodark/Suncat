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
    }
}
