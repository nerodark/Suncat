using System;
using System.ComponentModel;

namespace Cassia.Impl
{
    /// <summary>
    /// Connection to a remote terminal server.
    /// </summary>
    public class RemoteServerHandle : ITerminalServerHandle
    {
        private const string _connectionClosedMessage =
            "Connection to remote server not open. Did you: (1) forget to call ITerminalServer.Open, " +
            "or (2) did you try to perform operations on a session or process after closing the " +
            "connection to the server?";

        private readonly string _serverName;
        private IntPtr _serverPtr;

        public RemoteServerHandle(string serverName)
        {
            if (serverName == null)
            {
                throw new ArgumentNullException("serverName");
            }
            _serverName = serverName;
        }

        #region ITerminalServerHandle Members

        public IntPtr Handle
        {
            get
            {
                if (!IsOpen)
                {
                    throw new InvalidOperationException(_connectionClosedMessage);
                }
                return _serverPtr;
            }
        }

        public string ServerName
        {
            get { return _serverName; }
        }

        public bool IsOpen
        {
            get { return _serverPtr != IntPtr.Zero; }
        }

        public void Open()
        {
            if (_serverPtr != IntPtr.Zero)
            {
                return;
            }
            _serverPtr = NativeMethods.WTSOpenServer(_serverName);
            if (_serverPtr == IntPtr.Zero)
            {
                // Failed to connect, possibly because Terminal Services is not running on the remote machine.
                throw new Win32Exception();
            }
        }

        public void Close()
        {
            if (_serverPtr == IntPtr.Zero)
            {
                return;
            }
            NativeMethods.WTSCloseServer(_serverPtr);
            _serverPtr = IntPtr.Zero;
        }

        public bool Local
        {
            get { return false; }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion

        protected void Dispose(bool disposing)
        {
            Close();
        }

        ~RemoteServerHandle()
        {
            Dispose(false);
        }
    }
}