using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Windows.Forms;

namespace SuncatService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        //[DllImport("user32")]
        //private static extern bool ExitWindowsEx(uint uFlags, uint dwReason);

        private void ServiceInstaller_Committed(object sender, InstallEventArgs e)
        {
            Directory.SetCurrentDirectory(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

            //using (var process = Process.Start(new ProcessStartInfo()
            //{
            //    FileName = "secedit",
            //    Arguments = "/configure /db C:\\Windows\\System32\\secedit.sdb /cfg \"suncatprivs.inf\" /overwrite /quiet",
            //    WindowStyle = ProcessWindowStyle.Hidden,
            //    UseShellExecute = false,
            //    CreateNoWindow = true,
            //}))
            //{
            //    process.WaitForExit();
            //}
            
            // Auto-start service after install
            using (var sc = new ServiceController(ServiceInstaller.ServiceName))
            {
                sc.Start();
            }

            //if (MessageBox.Show(
            //    "Group policy settings have changed.\nYou need to log off for changes to take effect.\n\nLog off now?",
            //    "Log off",
            //    MessageBoxButtons.YesNoCancel,
            //    MessageBoxIcon.Question,
            //    MessageBoxDefaultButton.Button1,
            //    MessageBoxOptions.ServiceNotification) == DialogResult.Yes)
            //{
            //    // logoff to refresh user rights group policy
            //    try
            //    {
            //        ExitWindowsEx(0, 0);
            //    }
            //    catch (Exception ex)
            //    {
            //        Debug.WriteLine(ex);

            //        if (ex.InnerException != null)
            //            Debug.WriteLine(ex.InnerException);
            //    }
            //}
        }
    }
}
