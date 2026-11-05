using System.Text.Json;
using MailTemplater.Models;

namespace MailTemplater.Services;

public class TemplateStore
{
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };
    public string Root { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailTemplater");
    private string TemplatesFile => Path.Combine(Root, "templates.json");
    private string SettingsFile => Path.Combine(Root, "settings.json");

    public TemplateStore() => Directory.CreateDirectory(Root);

    public List<MailTemplate> Load()
    {
        try
        {
            if (File.Exists(TemplatesFile))
                return JsonSerializer.Deserialize<List<MailTemplate>>(File.ReadAllText(TemplatesFile)) ?? new();
        }
        catch { /* bozuk dosya: boş liste ile devam */ }
        return new();
    }

    public void Save(List<MailTemplate> list) =>
        File.WriteAllText(TemplatesFile, JsonSerializer.Serialize(list, Json));

    public AppSettings LoadSettings()
    {
        try
        {
            if (File.Exists(SettingsFile))
                return JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsFile)) ?? new();
        }
        catch { }
        return new();
    }

    public void SaveSettings(AppSettings s) =>
        File.WriteAllText(SettingsFile, JsonSerializer.Serialize(s, Json));

    public string AttachmentDir(MailTemplate t)
    {
        var dir = Path.Combine(Root, "Attachments", t.Id.ToString());
        Directory.CreateDirectory(dir);
        return dir;
    }

    public string AttachmentPath(MailTemplate t, string name) => Path.Combine(AttachmentDir(t), name);

    public void AddAttachmentFile(MailTemplate t, string sourcePath)
    {
        var name = Path.GetFileName(sourcePath);
        File.Copy(sourcePath, AttachmentPath(t, name), overwrite: true);
        if (!t.Attachments.Contains(name)) t.Attachments.Add(name);
    }

    public void AddAttachmentBytes(MailTemplate t, string fileName, byte[] data)
    {
        foreach (var c in Path.GetInvalidFileNameChars()) fileName = fileName.Replace(c, '_');
        if (string.IsNullOrWhiteSpace(fileName)) fileName = "ek";
        File.WriteAllBytes(AttachmentPath(t, fileName), data);
        if (!t.Attachments.Contains(fileName)) t.Attachments.Add(fileName);
    }

    public void RemoveAttachment(MailTemplate t, string name)
    {
        try { File.Delete(AttachmentPath(t, name)); } catch { }
        t.Attachments.Remove(name);
    }

    public void DeleteTemplateFiles(MailTemplate t)
    {
        try { Directory.Delete(Path.Combine(Root, "Attachments", t.Id.ToString()), true); } catch { }
    }
}
