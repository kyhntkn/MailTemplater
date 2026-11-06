using System.Runtime.InteropServices;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using MailTemplater.Models;

namespace MailTemplater.Services;

public class TemplateStore
{
    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping // Türkçe karakterler \u011F diye değil olduğu gibi yazılır
    };
    private static readonly UTF8Encoding Utf8 = new(false);

    public string Root { get; } = ResolveRoot();
    private string TemplatesFile => Path.Combine(Root, "templates.json");
    private string SettingsFile => Path.Combine(Root, "settings.json");

    public TemplateStore() => Directory.CreateDirectory(Root);

    // MSIX (Store) paketinde %LOCALAPPDATA% yazmaları pakete özel yere yönlendirilir ve Outlook o dosyaları göremez.
    // Bu yüzden paketliyken paketin gerçek (yönlendirilmemiş) LocalState klasörünü kullanıyoruz.
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetCurrentPackageFullName(ref int length, StringBuilder? name);

    private static string ResolveRoot()
    {
        try
        {
            int len = 0;
            if (GetCurrentPackageFullName(ref len, null) == 122) // ERROR_INSUFFICIENT_BUFFER => paketli
                return Path.Combine(Windows.Storage.ApplicationData.Current.LocalFolder.Path, "MailTemplater");
        }
        catch { /* paketsiz çalışıyor */ }
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailTemplater");
    }

    public List<MailTemplate> Load()
    {
        try
        {
            if (File.Exists(TemplatesFile))
            {
                var list = JsonSerializer.Deserialize<List<MailTemplate>>(File.ReadAllText(TemplatesFile, Utf8)) ?? new();
                foreach (var t in list) t.Attachments = t.Attachments.Select(TextFix.Clean).ToList();
                return list;
            }
        }
        catch { /* bozuk dosya: boş liste ile devam */ }
        return new();
    }

    public void Save(List<MailTemplate> list) =>
        File.WriteAllText(TemplatesFile, JsonSerializer.Serialize(list, Json), Utf8);

    public AppSettings LoadSettings()
    {
        try
        {
            if (File.Exists(SettingsFile))
                return JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsFile, Utf8)) ?? new();
        }
        catch { }
        return new();
    }

    public void SaveSettings(AppSettings s) =>
        File.WriteAllText(SettingsFile, JsonSerializer.Serialize(s, Json), Utf8);

    public string AttachmentDir(MailTemplate t)
    {
        var dir = Path.Combine(Root, "Attachments", t.Id.ToString());
        Directory.CreateDirectory(dir);
        return dir;
    }

    public string AttachmentPath(MailTemplate t, string name) => Path.Combine(AttachmentDir(t), name);

    private static string SafeName(string? name)
    {
        var n = TextFix.Clean(name);
        foreach (var c in Path.GetInvalidFileNameChars()) n = n.Replace(c, '_');
        return string.IsNullOrWhiteSpace(n) ? "ek" : n.Trim();
    }

    public void AddAttachmentFile(MailTemplate t, string sourcePath)
    {
        var name = SafeName(Path.GetFileName(sourcePath));
        File.Copy(sourcePath, AttachmentPath(t, name), overwrite: true);
        if (!t.Attachments.Contains(name)) t.Attachments.Add(name);
    }

    public void AddAttachmentBytes(MailTemplate t, string? fileName, byte[] data)
    {
        var name = SafeName(fileName);
        File.WriteAllBytes(AttachmentPath(t, name), data);
        if (!t.Attachments.Contains(name)) t.Attachments.Add(name);
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
