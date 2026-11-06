using System.Text;
using System.Windows;
namespace MailTemplater;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        // Türkçe (ISO-8859-9 / Windows-1254) gibi eski kod sayfalarını açar; .msg/.eml okuma için gerekli.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        base.OnStartup(e);
    }
}
