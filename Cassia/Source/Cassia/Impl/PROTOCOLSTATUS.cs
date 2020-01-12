using System.Runtime.InteropServices;

namespace Cassia.Impl
{
    [StructLayout(LayoutKind.Sequential)]
    public struct PROTOCOLSTATUS
    {
        public PROTOCOLCOUNTERS Output;
        public PROTOCOLCOUNTERS Input;
        public CACHE_STATISTICS Statistics;
        public int AsyncSignal;
        public int AsyncSignalMask;
    }
}