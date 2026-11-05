namespace MailTemplater.Models;

public class MailTemplate
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = "Yeni şablon";
    public string To { get; set; } = "";
    public string Cc { get; set; } = "";
    public string Bcc { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Body { get; set; } = "";
    public bool IsHtml { get; set; }
    /// <summary>Şablon klasörüne kopyalanmış ek dosya adları.</summary>
    public List<string> Attachments { get; set; } = new();
    public override string ToString() => Name;
}

public class AppSettings
{
    /// <summary>Yeni Outlook için Azure (Entra) uygulama Client ID'si.</summary>
    public string GraphClientId { get; set; } = "";
}
