using System.Net;
using System.Text.RegularExpressions;

namespace MailTemplater.Services;

public static class HtmlToText
{
    private const RegexOptions I = RegexOptions.IgnoreCase;

    /// <summary>Önizleme/düzenleme için güvenlik: script ve on*= olay özniteliklerini kaldırır.</summary>
    public static string StripScripts(string html)
    {
        html = Regex.Replace(html, @"<script\b[^>]*>.*?</script\s*>", "", I | RegexOptions.Singleline);
        html = Regex.Replace(html, @"\son\w+\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)", "", I);
        return html;
    }

    public static string Convert(string html)
    {
        if (string.IsNullOrEmpty(html)) return "";
        var s = Regex.Replace(html, @"<(script|style|head)[^>]*>.*?</\1>", "", I | RegexOptions.Singleline);
        s = Regex.Replace(s, @"<!--.*?-->", "", RegexOptions.Singleline);
        s = Regex.Replace(s, @"<\s*br\s*/?>", "\n", I);
        s = Regex.Replace(s, @"<\s*li[^>]*>", "• ", I);
        s = Regex.Replace(s, @"</\s*(p|div|tr|h[1-6]|li|table|blockquote)\s*>", "\n", I);
        s = Regex.Replace(s, @"</\s*t[dh]\s*>", "\t", I);
        s = Regex.Replace(s, @"<[^>]+>", "");
        s = WebUtility.HtmlDecode(s).Replace("\u00a0", " ").Replace("\r", "");
        s = Regex.Replace(s, @"[ \t]+\n", "\n");
        s = Regex.Replace(s, @"\n{3,}", "\n\n");
        return s.Trim();
    }
}
