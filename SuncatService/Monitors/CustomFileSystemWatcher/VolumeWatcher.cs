using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Threading;
using System.Windows.Forms;

namespace SuncatService.Monitors.CustomFileSystemWatcher
{
    public class VolumeWatcher : NativeWindow
    {
        private const int WM_DEVICECHANGE = 0x0219;
        private const int DBT_DEVICEARRIVAL = 0x8000;
        private const int DBT_DEVICEREMOVECOMPLETE = 0x8004;
        private const int DBT_DEVTYP_VOLUME = 0x00000002;
        //private const uint DRIVE_REMOVABLE = 2;

        #pragma warning disable 0649
        [StructLayout(LayoutKind.Sequential)]
        private struct DEV_BROADCAST_VOLUME
        {
            public uint dbch_Size;
            public uint dbch_Devicetype;
            public uint dbch_Reserved;
            public uint dbch_Unitmask;
            public ushort dbch_Flags;
        }
        #pragma warning restore 0649

        private ServiceController serviceController;
        private ManualResetEvent serviceEvent;
        private ManagementEventWatcher volumeWatcher;
        private WqlEventQuery volumeQuery;
        private Thread serviceChecker;
        private bool serviceIsRunning;

        public Dictionary<string, VolumeInfo> DriveList { get; private set; }
        private Thread driveChecker;
        private object driveLock = new object();

        private Form parent;

        public event EventHandler<VolumeChangedEventArgs> DriveInserted;
        public event EventHandler<VolumeChangedEventArgs> DriveRemoved;

        public VolumeWatcher(Form parent = null)
        {
            //SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;

            if (parent != null)
            {
                parent.HandleCreated += Parent_HandleCreated;
                parent.HandleDestroyed += Parent_HandleDestroyed;
                this.parent = parent;
            }
        }

        public void Start()
        {
            serviceController = new ServiceController("Winmgmt");
            serviceEvent = new ManualResetEvent(true);

            volumeQuery = new WqlEventQuery();
            volumeQuery.EventClassName = "Win32_VolumeChangeEvent";
            volumeQuery.Condition = "EventType = 2 OR EventType = 3";

            serviceChecker = new Thread(() => CheckService());
            serviceChecker.IsBackground = true;
            serviceChecker.Start();

            DriveList = new Dictionary<string, VolumeInfo>();

            for (int i = 0; i < 26; i++)
            {
                var driveLetter = (char)(i + 'A');

                DriveList.Add(driveLetter.ToString(), new VolumeInfo() { DriveType = DriveType.Unknown, EventType = EventType.Initial });
            }

            driveChecker = new Thread(() => CheckDrives());
            driveChecker.IsBackground = true;
            driveChecker.Start();
        }

        //private void SystemEvents_PowerModeChanged(object sender, PowerModeChangedEventArgs e)
        //{
        //    switch (e.Mode)
        //    {
        //        case PowerModes.Suspend:
        //            break;

        //        case PowerModes.Resume:
        //            break;
        //    }
        //}

        private void Parent_HandleCreated(object sender, EventArgs e)
        {
            this.AssignHandle(((Form)sender).Handle);
        }

        private void Parent_HandleDestroyed(object sender, EventArgs e)
        {
            this.ReleaseHandle();
        }

        protected override void WndProc(ref Message m)
        {
            if (!serviceIsRunning)
            {
                switch (m.Msg)
                {
                    case WM_DEVICECHANGE:
                        if (m.LParam != IntPtr.Zero)
                        {
                            if (Marshal.ReadInt32(m.LParam, 4) == DBT_DEVTYP_VOLUME)
                            {
                                DEV_BROADCAST_VOLUME volume = (DEV_BROADCAST_VOLUME)Marshal.PtrToStructure(m.LParam, typeof(DEV_BROADCAST_VOLUME));
                                string driveName;

                                switch (m.WParam.ToInt32())
                                {
                                    case DBT_DEVICEARRIVAL:
                                        driveName = FirstDriveFromMask(volume.dbch_Unitmask);

                                        lock (driveLock)
                                        {
                                            var driveInfo = new DriveInfo(driveName.ToString());

                                            if (IsDriveTypeSupported(driveInfo.DriveType) && driveInfo.IsReady && DriveList[driveName].EventType != EventType.Inserted)
                                            {
                                                if (DriveInserted != null)
                                                {
                                                    DriveInserted(null, new VolumeChangedEventArgs() { DriveLetter = driveName, DriveType = driveInfo.DriveType });
                                                }

                                                DriveList[driveName].DriveType = driveInfo.DriveType;
                                                DriveList[driveName].EventType = EventType.Inserted;
                                            }
                                        }
                                        break;

                                    case DBT_DEVICEREMOVECOMPLETE:
                                        driveName = FirstDriveFromMask(volume.dbch_Unitmask);

                                        lock (driveLock)
                                        {
                                            if (IsDriveTypeSupported(DriveList[driveName].DriveType) && DriveList[driveName].EventType != EventType.Initial && DriveList[driveName].EventType != EventType.Removed)
                                            {
                                                if (DriveRemoved != null)
                                                {
                                                    DriveRemoved(null, new VolumeChangedEventArgs() { DriveLetter = driveName, DriveType = DriveList[driveName].DriveType });
                                                }

                                                DriveList[driveName].DriveType = DriveType.Unknown;
                                                DriveList[driveName].EventType = EventType.Removed;
                                            }
                                        }
                                        break;
                                }
                            }
                        }
                        break;
                }
            }

            base.WndProc(ref m);
        }

        private void VolumeWatcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
            var driveName = ((string)e.NewEvent.Properties["DriveName"].Value).Substring(0, 1);
            var eventType = (EventType)Convert.ToInt16((e.NewEvent.Properties["EventType"].Value));

            switch (eventType)
            {
                case EventType.Inserted:
                    lock (driveLock)
                    {
                        var driveInfo = new DriveInfo(driveName.ToString());

                        if (IsDriveTypeSupported(driveInfo.DriveType) && driveInfo.IsReady && DriveList[driveName].EventType != EventType.Inserted)
                        {
                            if (DriveInserted != null)
                            {
                                DriveInserted(null, new VolumeChangedEventArgs() { DriveLetter = driveName, DriveType = driveInfo.DriveType });
                            }

                            DriveList[driveName].DriveType = driveInfo.DriveType;
                            DriveList[driveName].EventType = EventType.Inserted;
                        }
                    }
                    break;

                case EventType.Removed:
                    lock (driveLock)
                    {
                        if (IsDriveTypeSupported(DriveList[driveName].DriveType) && DriveList[driveName].EventType != EventType.Initial && DriveList[driveName].EventType != EventType.Removed)
                        {
                            if (DriveRemoved != null)
                            {
                                DriveRemoved(null, new VolumeChangedEventArgs() { DriveLetter = driveName, DriveType = DriveList[driveName].DriveType });
                            }

                            DriveList[driveName].DriveType = DriveType.Unknown;
                            DriveList[driveName].EventType = EventType.Removed;
                        }
                    }
                    break;
            }
        }

        private void VolumeWatcher_Stopped(object sender, StoppedEventArgs e)
        {
            volumeWatcher.Dispose();

            serviceIsRunning = false;
            serviceEvent.Set();
        }

        private void StartVolumeWatcher()
        {
            volumeWatcher = new ManagementEventWatcher();

            volumeWatcher.EventArrived += VolumeWatcher_EventArrived;
            volumeWatcher.Stopped += VolumeWatcher_Stopped;
            volumeWatcher.Query = volumeQuery;

            volumeWatcher.Start();
        }

        private void CheckService()
        {
            serviceIsRunning = serviceController.Status == ServiceControllerStatus.Running;

            while (true)
            {
                serviceEvent.WaitOne();

                if (serviceIsRunning)
                {
                    StartVolumeWatcher();
                    serviceEvent.Reset();
                }
                else
                {
                    serviceController.WaitForStatus(ServiceControllerStatus.Running);
                    serviceIsRunning = true;
                }
            }
        }

        private string FirstDriveFromMask(uint unitmask)
        {
            int i;

            for (i = 0; i < 26; i++)
            {
                if ((unitmask & 0x1) == 0x1)
                {
                    break;
                }

                unitmask = unitmask >> 1;
            }

            var driveLetter = (char)(i + 'A');
            return driveLetter.ToString();
        }

        private bool IsDriveTypeSupported(DriveType driveType)
        {
            switch (driveType)
            {
                case DriveType.CDRom:
                case DriveType.Fixed:
                case DriveType.Network:
                case DriveType.Ram:
                case DriveType.Removable:
                    return true;

                default:
                    return false;
            }
        }

        private void CheckDrives()
        {
            while (true)
            {
                var drives = DriveInfo.GetDrives();

                lock (driveLock)
                {
                    foreach (var drive in DriveList)
                    {
                        var driveInfo = drives.SingleOrDefault(d => d.Name.Substring(0, 1) == drive.Key);

                        if (driveInfo != null)
                        {
                            if (IsDriveTypeSupported(driveInfo.DriveType) && driveInfo.IsReady && drive.Value.EventType != EventType.Inserted)
                            {
                                if (DriveInserted != null)
                                {
                                    DriveInserted(null, new VolumeChangedEventArgs() { DriveLetter = drive.Key, DriveType = driveInfo.DriveType });
                                }

                                drive.Value.DriveType = driveInfo.DriveType;
                                drive.Value.EventType = EventType.Inserted;
                            }
                            else if (IsDriveTypeSupported(drive.Value.DriveType) && !driveInfo.IsReady && drive.Value.EventType != EventType.Initial && drive.Value.EventType != EventType.Removed)
                            {
                                if (DriveRemoved != null)
                                {
                                    DriveRemoved(null, new VolumeChangedEventArgs() { DriveLetter = drive.Key, DriveType = drive.Value.DriveType });
                                }

                                drive.Value.DriveType = DriveType.Unknown;
                                drive.Value.EventType = EventType.Removed;
                            }
                        }
                        else if (drive.Value.EventType != EventType.Initial && drive.Value.EventType != EventType.Removed)
                        {
                            if (IsDriveTypeSupported(drive.Value.DriveType))
                            {
                                driveInfo = new DriveInfo(drive.Key.ToString());

                                if (!driveInfo.IsReady)
                                {
                                    if (DriveRemoved != null)
                                    {
                                        DriveRemoved(null, new VolumeChangedEventArgs() { DriveLetter = drive.Key, DriveType = drive.Value.DriveType });
                                    }

                                    drive.Value.DriveType = DriveType.Unknown;
                                    drive.Value.EventType = EventType.Removed;
                                }
                            }
                        }
                    }
                }

                Thread.Sleep(1000);
            }
        }
    }

    public enum EventType { Initial, Inserted = 2, Removed }

    public class VolumeInfo
    {
        public DriveType DriveType { get; set; }
        public EventType EventType { get; set; }
    }

    public class VolumeChangedEventArgs : EventArgs
    {
        public string DriveLetter { get; set; }
        public DriveType DriveType { get; set; }
    }
}
