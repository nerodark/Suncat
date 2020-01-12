using System;
using System.IO;

namespace SuncatObjects
{
    public enum SuncatLogEvent
    {
        None,
        CreateFile,
        DeleteFile,
        ChangeFile,
        MoveFile,
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
}
