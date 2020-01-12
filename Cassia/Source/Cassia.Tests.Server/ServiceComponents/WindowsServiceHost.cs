using System.ServiceProcess;

namespace Cassia.Tests.Server.ServiceComponents
{
    public class WindowsServiceHost : ServiceBase, IServiceHost
    {
        private IHostedService _service;

        #region IServiceHost Members

        public void Run(IHostedService service)
        {
            _service = service;
            ServiceName = _service.Name;
            Logger.SetLogger(new EventLogLogger(EventLog));
            _service.Attach(this);
            Run(this);
        }

        #endregion

        protected override void OnStart(string[] args)
        {
            _service.Start();
        }

        protected override void OnStop()
        {
            _service.Stop();
        }
    }
}