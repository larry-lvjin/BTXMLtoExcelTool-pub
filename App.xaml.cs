using System.Windows;
using BtFiller2.Services;

namespace BtFiller2
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            LicenseGuard.EnforceOrExit();
            base.OnStartup(e);
        }
    }
}
