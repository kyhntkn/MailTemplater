using MailTemplater.Models;
using MsgReader.Outlook;

namespace MailTemplater.Services;

/// <summary>.msg ve .eml dosyalarından To/CC/BCC/konu/içerik/ek okur.</summary>
public static class MailImporter
{
    public static MailTemplate Import(string path, TemplateStore store)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        var t = new MailTemplate { Name = Path.GetFileNameWithoutExtension(path) };
        if (ext == ".msg") ReadMsg(path, t, store);
        else if (ext == ".eml") ReadEml(path, t, store);
        else throw new NotSupportedException("Yalnızca .msg ve .eml desteklenir.");
        if (!string.IsNullOrWhiteSpace(t.Subject)) t.Name = t.Subject;
        return t;
    }

    private static void ReadMsg(string path, MailTemplate t, TemplateStore store)
    {
        using var msg = new Storage.Message(path);
        t.Subject = msg.Subject ?? "";

        string Join(RecipientType type) => string.Join("; ",
            msg.Recipients.Where(r => r.Type == type)
               .Select(r => string.IsNullOrWhiteSpace(r.Email) ? r.DisplayName : r.Email)
               .Where(s => !string.IsNullOrWhiteSpace(s)));

        t.To = Join(RecipientType.To);
        t.Cc = Join(RecipientType.Cc);
        t.Bcc = Join(RecipientType.Bcc);

        if (!string.IsNullOrWhiteSpace(msg.BodyHtml)) { t.Body = msg.BodyHtml; t.IsHtml = true; }
        else t.Body = msg.BodyText ?? "";

        foreach (var o in msg.Attachments)
            if (o is Storage.Attachment a && !a.IsInline && a.Data != null)
                store.AddAttachmentBytes(t, a.FileName ?? "ek", a.Data);
    }

    private static void ReadEml(string path, MailTemplate t, TemplateStore store)
    {
        var m = MsgReader.Mime.Message.Load(new FileInfo(path));
        string J(IEnumerable<MsgReader.Mime.Header.RfcMailAddress>? l) =>
            l == null ? "" : string.Join("; ", l.Select(a => a.Address).Where(s => !string.IsNullOrWhiteSpace(s)));

        t.Subject = m.Headers.Subject ?? "";
        t.To = J(m.Headers.To);
        t.Cc = J(m.Headers.Cc);
        t.Bcc = J(m.Headers.Bcc);

        var html = m.HtmlBody?.GetBodyAsText();
        if (!string.IsNullOrWhiteSpace(html)) { t.Body = html; t.IsHtml = true; }
        else t.Body = m.TextBody?.GetBodyAsText() ?? "";

        foreach (var a in m.Attachments)
            if (a.Body != null)
                store.AddAttachmentBytes(t, a.FileName ?? "ek", a.Body);
    }
}
