using System;
using System.IO;

namespace SuncatCommon
{
    public enum SuncatLogEvent
    {
        None,
        CreateFile,
        DeleteFile,
        ChangeFile,
        RenameFile,
        OpenFile,
        CopyFile,
        OpenURL,
        SwitchApp,
        TypeOnKeyboard,
        CopyToClipboard,
        HeartBeat,
        GetPublicIP,
        Explorer,
    }

    public class TrackEventArgs : EventArgs
    {
        public SuncatLog LogData { get; set; }
    }

    public delegate void TrackEventHandler(object sender, TrackEventArgs e);

    [Serializable]
    public class SuncatLog
    {
        public SuncatLog()
        { 
        }

        public SuncatLog(FileSystemEventArgs e)
        {
            Event = e.ChangeType.ToSuncatLogEvent();

            if (Event == SuncatLogEvent.RenameFile)
            {
                var rea = (RenamedEventArgs)e;
                Data1 = rea.OldFullPath;
                Data2 = rea.FullPath;
            }
            else
            {
                Data1 = e.FullPath;
            }
        }

        public DateTime DateTime { get; set; }
        public SuncatLogEvent Event { get; set; }
        public object DataObject { get; set; }
        public string Data1 { get; set; }
        public string Data2 { get; set; }
        public string Data3 { get; set; }
    }

    [Serializable]
    public class SuncatFileInfo
    {
        public FileInfo FileInfo { get; set; }
        public bool IsDirectory { get; set; }
    }

    public static class SuncatExtensions
    {
        public static SuncatLogEvent ToSuncatLogEvent(this WatcherChangeTypes changeType)
        {
            SuncatLogEvent logEvent = SuncatLogEvent.None;

            switch (changeType)
            {
                case WatcherChangeTypes.Created: logEvent = SuncatLogEvent.CreateFile; break;
                case WatcherChangeTypes.Deleted: logEvent = SuncatLogEvent.DeleteFile; break;
                case WatcherChangeTypes.Changed: logEvent = SuncatLogEvent.ChangeFile; break;
                case WatcherChangeTypes.Renamed: logEvent = SuncatLogEvent.RenameFile; break;
            }

            return logEvent;
        }
    }
}
