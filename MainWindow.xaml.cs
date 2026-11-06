using System.Net;
using System.Text;
using System.Text.Json;
using Microsoft.Web.WebView2.Core;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using MailTemplater.Models;
using MailTemplater.Services;
using Microsoft.Win32;

namespace MailTemplater;

public partial class MainWindow : Window
{
    private readonly TemplateStore _store = new();
    private List<MailTemplate> _templates;
    private MailTemplate? _cur;
    private bool _loading;
    private bool _webViewOk;
    private bool _previewDirty;

    public MainWindow()
    {
        InitializeComponent();
        _templates = _store.Load();
        ClientIdBox.Text = _store.LoadSettings().GraphClientId;
        RefreshList();
        Loaded += async (_, _) => await InitPreviewAsync();
    }

    // ---------- Önizleme ----------
    private async Task InitPreviewAsync()
    {
        try
        {
            var env = await CoreWebView2Environment.CreateAsync(null, Path.Combine(_store.Root, "WebView2"));
            await Preview.EnsureCoreWebView2Async(env);
            // Düzenleme için script gerekir; mail içeriğindeki script'ler UpdatePreview'da temizlenir.
            Preview.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            Preview.CoreWebView2.Settings.AreDevToolsEnabled = false;
            Preview.CoreWebView2.WebMessageReceived += (_, _) => _previewDirty = true;
            Preview.CoreWebView2.NavigationCompleted += async (_, _) =>
            {
                try
                {
                    await Preview.CoreWebView2.ExecuteScriptAsync(
                        "document.designMode='on';document.addEventListener('input',function(){window.chrome.webview.postMessage('d');});");
                }
                catch { }
            };
            Preview.CoreWebView2.NavigationStarting += (_, e) =>
            { if (!e.Uri.StartsWith("data:") && e.Uri != "about:blank") e.Cancel = true; };
            Preview.CoreWebView2.NewWindowRequested += (_, e) => e.Handled = true;
            _webViewOk = true;
            UpdatePreview();
        }
        catch { _webViewOk = false; BodyTabs.SelectedIndex = 1; }
    }

    private void UpdatePreview()
    {
        if (!_webViewOk) return;
        try
        {
            var html = HtmlCheck.IsChecked == true
                ? BodyBox.Text
                : "<pre style=\"font-family:Segoe UI,sans-serif;white-space:pre-wrap\">" + WebUtility.HtmlEncode(BodyBox.Text) + "</pre>";
            _previewDirty = false;
            Preview.NavigateToString(string.IsNullOrEmpty(html) ? "<html></html>" : HtmlToText.StripScripts(html));
        }
        catch { }
    }

    private async void BodyTabs_SelectionChanged(object s, SelectionChangedEventArgs e)
    {
        if (e.Source != BodyTabs) return;
        if (BodyTabs.SelectedIndex == 0) UpdatePreview();
        else await SyncFromPreviewAsync();
    }

    /// <summary>Önizlemede yapılan düzenlemeyi içerik kutusuna geri yazar.</summary>
    private async Task SyncFromPreviewAsync()
    {
        if (!_webViewOk || !_previewDirty) return;
        try
        {
            bool html = HtmlCheck.IsChecked == true;
            var json = await Preview.CoreWebView2.ExecuteScriptAsync(
                html ? "document.documentElement.outerHTML" : "document.body.innerText");
            var text = JsonSerializer.Deserialize<string>(json) ?? "";
            BodyBox.Text = html ? "<!DOCTYPE html>\n" + text : text;
            _previewDirty = false;
        }
        catch { }
    }

    private async Task PrepareAsync()
    {
        await SyncFromPreviewAsync();
        FormToModel();
    }

    private async Task FormatAsync(string cmd)
    {
        if (!_webViewOk || BodyTabs.SelectedIndex != 0) return;
        if (HtmlCheck.IsChecked != true)
        {
            MessageBox.Show("Biçimlendirme için önce 'İçerik HTML' kutusunu işaretleyin.");
            return;
        }
        await Preview.CoreWebView2.ExecuteScriptAsync($"document.execCommand('{cmd}')");
        _previewDirty = true;
    }

    private async void Bold_Click(object s, RoutedEventArgs e) => await FormatAsync("bold");
    private async void Italic_Click(object s, RoutedEventArgs e) => await FormatAsync("italic");
    private async void Underline_Click(object s, RoutedEventArgs e) => await FormatAsync("underline");

    private async void ToPlain_Click(object s, RoutedEventArgs e)
    {
        await SyncFromPreviewAsync();
        if (HtmlCheck.IsChecked == true &&
            MessageBox.Show("HTML biçimlendirme (renk, tablo, resim vb.) kaybolur ve içerik düz metin olur. Devam edilsin mi?",
                            "Onay", MessageBoxButton.YesNo) != MessageBoxResult.Yes) return;
        BodyBox.Text = HtmlToText.Convert(BodyBox.Text);
        HtmlCheck.IsChecked = false;
        BodyTabs.SelectedIndex = 1;
        UpdatePreview();
    }

    // ---------- Liste / form ----------
    private void RefreshList(MailTemplate? select = null)
    {
        _loading = true;
        TemplateList.ItemsSource = null;
        TemplateList.ItemsSource = _templates;
        TemplateList.SelectedItem = select;
        _loading = false;
        LoadToForm(select);
    }

    private void LoadToForm(MailTemplate? t)
    {
        _cur = t;
        _previewDirty = false;
        NameBox.Text = t?.Name ?? ""; ToBox.Text = t?.To ?? ""; CcBox.Text = t?.Cc ?? "";
        BccBox.Text = t?.Bcc ?? ""; SubjectBox.Text = t?.Subject ?? ""; BodyBox.Text = t?.Body ?? "";
        HtmlCheck.IsChecked = t?.IsHtml ?? false;
        AttachList.ItemsSource = null;
        AttachList.ItemsSource = t?.Attachments;
        BodyTabs.SelectedIndex = (t?.IsHtml ?? false) ? 0 : 1;
        UpdatePreview();
    }

    private void FormToModel()
    {
        if (_cur == null) return;
        _cur.Name = string.IsNullOrWhiteSpace(NameBox.Text) ? "Adsız şablon" : NameBox.Text.Trim();
        _cur.To = ToBox.Text.Trim(); _cur.Cc = CcBox.Text.Trim(); _cur.Bcc = BccBox.Text.Trim();
        _cur.Subject = SubjectBox.Text; _cur.Body = BodyBox.Text; _cur.IsHtml = HtmlCheck.IsChecked == true;
    }

    private void TemplateList_SelectionChanged(object s, SelectionChangedEventArgs e)
    {
        if (_loading) return;
        LoadToForm(TemplateList.SelectedItem as MailTemplate);
    }

    private void New_Click(object s, RoutedEventArgs e)
    {
        var t = new MailTemplate();
        _templates.Add(t);
        _store.Save(_templates);
        RefreshList(t);
    }

    private async void Save_Click(object s, RoutedEventArgs e)
    {
        if (_cur == null) { MessageBox.Show("Önce bir şablon seçin veya oluşturun."); return; }
        await PrepareAsync();
        _store.Save(_templates);
        RefreshList(_cur);
    }

    private void Delete_Click(object s, RoutedEventArgs e)
    {
        if (_cur == null) return;
        if (MessageBox.Show($"'{_cur.Name}' silinsin mi?", "Onay", MessageBoxButton.YesNo) != MessageBoxResult.Yes) return;
        _store.DeleteTemplateFiles(_cur);
        _templates.Remove(_cur);
        _store.Save(_templates);
        RefreshList();
    }

    // ---------- Ekler ----------
    private async void AddAttach_Click(object s, RoutedEventArgs e)
    {
        if (_cur == null) return;
        var dlg = new OpenFileDialog { Multiselect = true };
        if (dlg.ShowDialog() != true) return;
        foreach (var f in dlg.FileNames) _store.AddAttachmentFile(_cur, f);
        await PrepareAsync(); _store.Save(_templates); LoadToForm(_cur);
    }

    private async void RemoveAttach_Click(object s, RoutedEventArgs e)
    {
        if (_cur == null || AttachList.SelectedItem is not string name) return;
        _store.RemoveAttachment(_cur, name);
        await PrepareAsync(); _store.Save(_templates); LoadToForm(_cur);
    }

    // ---------- Sürükle-bırak ----------
    private void DropZone_DragEnter(object s, DragEventArgs e)
    {
        e.Effects = DragDropEffects.Copy;
        SetDropHighlight(true);
    }

    private void DropZone_DragOver(object s, DragEventArgs e)
    {
        e.Effects = DragDropEffects.Copy;
        e.Handled = true;
    }

    private void DropZone_DragLeave(object s, DragEventArgs e) => SetDropHighlight(false);

    private void SetDropHighlight(bool on)
    {
        DropZone.Background = (Brush)FindResource(on ? "DropActiveBrush" : "DropIdleBrush");
        DropOutline.Stroke = (Brush)FindResource(on ? "OutlineActiveBrush" : "OutlineIdleBrush");
    }

    private void DropZone_Drop(object s, DragEventArgs e)
    {
        SetDropHighlight(false);
        string? path = null;
        bool temp = false;
        try
        {
            // 1) Gerçek dosya (.msg/.eml masaüstünden vb.)
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                path = files.FirstOrDefault(f => f.EndsWith(".msg", StringComparison.OrdinalIgnoreCase)
                                              || f.EndsWith(".eml", StringComparison.OrdinalIgnoreCase));
            }
            // 2) Outlook'tan doğrudan sürükleme (sanal dosya)
            else if (e.Data.GetDataPresent("FileGroupDescriptorW"))
            {
                path = ExtractVirtualFile(e.Data);
                if (path == null) path = OutlookClassicLauncher.ExportSelectedMail(); // klasik Outlook yedeği
                temp = path != null;
            }

            if (path == null)
            {
                MessageBox.Show("Mail okunamadı. Maili önce masaüstüne sürükleyip (.msg/.eml) sonra buraya bırakmayı deneyin.\n" +
                                "Not: Outlook ile bu uygulama aynı yetkide (ikisi de normal kullanıcı) çalışmalı.");
                return;
            }

            var t = MailImporter.Import(path, _store);
            _templates.Add(t);
            _store.Save(_templates);
            RefreshList(t);
        }
        catch (Exception ex) { MessageBox.Show("İçe aktarma hatası: " + ex.Message); }
        finally { if (temp && path != null) try { File.Delete(path); } catch { } }
    }

    private static string? ExtractVirtualFile(IDataObject data)
    {
        try
        {
            if (data.GetData("FileGroupDescriptorW") is not MemoryStream desc) return null;
            var buf = desc.ToArray();
            // FILEGROUPDESCRIPTORW: 4 bayt sayaç + FILEDESCRIPTORW (cFileName ofset 72)
            var name = Encoding.Unicode.GetString(buf, 4 + 72, 520).Split('\0')[0];
            if (data.GetData("FileContents") is not MemoryStream content) return null;
            var ext = Path.GetExtension(name).ToLowerInvariant();
            if (ext != ".msg" && ext != ".eml") return null;
            var tmp = Path.Combine(Path.GetTempPath(), $"mt_{Guid.NewGuid():N}{ext}");
            File.WriteAllBytes(tmp, content.ToArray());
            return tmp;
        }
        catch { return null; }
    }

    // ---------- Outlook'ta açma ----------
    private async void OpenClassic_Click(object s, RoutedEventArgs e)
    {
        if (_cur == null) return;
        try { await PrepareAsync(); _store.Save(_templates); OutlookClassicLauncher.Open(_cur, _store); }
        catch (Exception ex) { MessageBox.Show(ex.Message); }
    }

    private async void OpenNew_Click(object s, RoutedEventArgs e)
    {
        if (_cur == null) return;
        try
        {
            await PrepareAsync(); _store.Save(_templates);
            await GraphLauncher.OpenAsync(_cur, _store, ClientIdBox.Text.Trim());
        }
        catch (Exception ex) { MessageBox.Show(ex.Message); }
    }

    private void ClientId_LostFocus(object s, RoutedEventArgs e) =>
        _store.SaveSettings(new AppSettings { GraphClientId = ClientIdBox.Text.Trim() });
}
