using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using MailTemplater.Models;
using Microsoft.Identity.Client;

namespace MailTemplater.Services;

/// <summary>
/// Yeni Outlook için: Microsoft Graph ile Taslaklar'da mail oluşturur ve açar.
/// Not: ~3 MB'a kadar toplam ek desteklenir (büyük ekler için upload session gerekir).
/// </summary>
public static class GraphLauncher
{
    private static readonly string[] Scopes = { "Mail.ReadWrite" };
    private static IPublicClientApplication? _app;
    private static string? _clientId;
    private static readonly HttpClient Http = new();

    public static async Task OpenAsync(MailTemplate t, TemplateStore store, string clientId)
    {
        if (string.IsNullOrWhiteSpace(clientId))
            throw new InvalidOperationException("Azure Client ID girilmedi (README'ye bakın).");

        if (_app == null || _clientId != clientId)
        {
            _clientId = clientId;
            _app = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AadAuthorityAudience.AzureAdAndPersonalMicrosoftAccount)
                .WithRedirectUri("http://localhost")
                .Build();
        }

        AuthenticationResult auth;
        var accounts = await _app.GetAccountsAsync();
        try { auth = await _app.AcquireTokenSilent(Scopes, accounts.FirstOrDefault()).ExecuteAsync(); }
        catch { auth = await _app.AcquireTokenInteractive(Scopes).ExecuteAsync(); }

        var attachments = new List<object>();
        long total = 0;
        foreach (var name in t.Attachments)
        {
            var p = store.AttachmentPath(t, name);
            if (!File.Exists(p)) continue;
            var bytes = await File.ReadAllBytesAsync(p);
            total += bytes.Length;
            attachments.Add(new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.fileAttachment",
                ["name"] = name,
                ["contentBytes"] = Convert.ToBase64String(bytes)
            });
        }
        if (total > 3 * 1024 * 1024)
            throw new InvalidOperationException("Toplam ek boyutu 3 MB'ı aşıyor; Graph yöntemi bu sürümde desteklemiyor.");

        static object[] Rcpt(string s) => s.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(a => (object)new { emailAddress = new { address = a } }).ToArray();

        var payload = new
        {
            subject = t.Subject,
            body = new { contentType = t.IsHtml ? "HTML" : "Text", content = t.Body },
            toRecipients = Rcpt(t.To),
            ccRecipients = Rcpt(t.Cc),
            bccRecipients = Rcpt(t.Bcc),
            attachments
        };

        using var req = new HttpRequestMessage(HttpMethod.Post, "https://graph.microsoft.com/v1.0/me/messages");
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", auth.AccessToken);
        req.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
        using var resp = await Http.SendAsync(req);
        var json = await resp.Content.ReadAsStringAsync();
        if (!resp.IsSuccessStatusCode) throw new InvalidOperationException($"Graph hatası: {(int)resp.StatusCode}\n{json}");

        var link = JsonDocument.Parse(json).RootElement.GetProperty("webLink").GetString();
        if (!string.IsNullOrEmpty(link))
            Process.Start(new ProcessStartInfo(link) { UseShellExecute = true });
    }
}
