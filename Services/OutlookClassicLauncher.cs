using MailTemplater.Models;

namespace MailTemplater.Services;

/// <summary>Klasik Outlook (COM). Yeni Outlook COM desteklemez.</summary>
public static class OutlookClassicLauncher
{
    private static dynamic GetApp()
    {
        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException(
                "Klasik Outlook bulunamadı (yeni Outlook COM desteklemez). 'Yeni Outlook'ta aç' seçeneğini kullanın.");
        return Activator.CreateInstance(type)!;
    }

    public static void Open(MailTemplate t, TemplateStore store)
    {
        dynamic app = GetApp();
        dynamic mail = app.CreateItem(0); // olMailItem
        mail.To = t.To;
        mail.CC = t.Cc;
        mail.BCC = t.Bcc;
        mail.Subject = t.Subject;
        foreach (var name in t.Attachments)
        {
            var p = store.AttachmentPath(t, name);
            if (File.Exists(p)) mail.Attachments.Add(p);
        }
        mail.Display(false);
        // İmza korunsun diye: önce pencere açılır, sonra içerik imzanın üstüne eklenir.
        if (t.IsHtml) mail.HTMLBody = t.Body + (string)mail.HTMLBody;
        else mail.Body = t.Body + "\r\n" + (string)mail.Body;
    }

    /// <summary>Outlook'tan sürüklenen mail sanal dosya olduğunda: seçili maili .msg olarak kaydeder.</summary>
    public static string? ExportSelectedMail()
    {
        try
        {
            dynamic app = GetApp();
            dynamic sel = app.ActiveExplorer().Selection;
            if (sel.Count < 1) return null;
            var tmp = Path.Combine(Path.GetTempPath(), $"mt_{Guid.NewGuid():N}.msg");
            sel[1].SaveAs(tmp, 3); // olMSG
            return File.Exists(tmp) ? tmp : null;
        }
        catch { return null; }
    }
}
