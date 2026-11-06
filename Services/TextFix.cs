using System.Text;
using System.Text.RegularExpressions;

namespace MailTemplater.Services;

/// <summary>Türkçe karakter sorunlarına karşı yardımcılar.</summary>
public static class TextFix
{
    /// <summary>Dosya adı vb. için: bozuk (mojibake) kodlamayı düzeltir ve Unicode'u FormC'ye normalleştirir.</summary>
    public static string Clean(string? s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        s = FixMojibake(s);
        return s.Normalize(NormalizationForm.FormC);
    }

    /// <summary>"ÅŸ", "Ã¼", "Ä±" gibi UTF-8'in yanlış okunmasından doğan bozuk metni geri çevirir.</summary>
    public static string FixMojibake(string s)
    {
        if (!s.Any(c => c == 'Ã' || c == 'Ä' || c == 'Å')) return s;
        try
        {
            var bytes = s.All(c => c <= 255) ? Encoding.Latin1.GetBytes(s) : Encoding.GetEncoding(1252).GetBytes(s);
            return new UTF8Encoding(false, true).GetString(bytes); // geçersizse hata verir, orijinal kalır
        }
        catch { return s; }
    }

    /// <summary>HTML içindeki eski charset bildirimlerini (windows-1254, iso-8859-9...) utf-8 yapar.
    /// İçerik zaten Unicode metne çevrildiği için eski bildirim Outlook'ta bozulmaya yol açabilir.</summary>
    public static string NormalizeHtmlCharset(string html)
    {
        if (string.IsNullOrEmpty(html)) return html;
        return Regex.Replace(html, @"(<meta[^>]*charset\s*=\s*[""']?)[\w\-]+", "${1}utf-8", RegexOptions.IgnoreCase);
    }
}
