# Mail Şablonları (WPF, .NET 8)

## Derleme / .exe
Gereken: Windows 10/11, Visual Studio 2022 (".NET masaüstü geliştirme") veya .NET 8 SDK.

    dotnet restore
    dotnet run
    dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
    # çıktı: bin\Release\net8.0-windows10.0.19041.0\win-x64\publish\MailTemplater.exe

## Microsoft Store (MSIX)
1. Visual Studio'da çözüme "Windows Application Packaging Project" ekleyin, Applications'a bu projeyi referans verin.
2. Package.appxmanifest: Capabilities -> `runFullTrust` (klasik Outlook COM için gerekli).
3. Proje sağ tık -> Publish -> Create App Packages -> Microsoft Store. Partner Center hesabı (tek seferlik bireysel ücret) gerekir.
4. Store ikonları (44, 150, 50 px vb.) Assets klasörüne eklenmeli.

## Yeni Outlook için Azure kaydı (ücretsiz)
1. portal.azure.com -> Microsoft Entra ID -> App registrations -> New registration
2. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
3. Redirect URI: platform "Mobile and desktop applications" -> `http://localhost`
4. API permissions -> Microsoft Graph -> Delegated -> `Mail.ReadWrite`
5. "Application (client) ID" değerini uygulamadaki Azure Client ID kutusuna yapıştırın.
Kurumsal hesapta yönetici onayı gerekebilir.

## Bilinen sınırlar
- Yeni Outlook COM desteklemez: Graph ile Taslaklar'da mail oluşturulur ve web bağlantısıyla açılır.
- Graph yönteminde toplam ek ~3 MB ile sınırlı.
- Gömülü (inline) resimler şablona alınmaz.
- Gönderilmiş maillerde BCC genelde dosyada bulunmaz.

## GitHub Actions ile derleme
`.github/workflows/build.yml` her push'ta `MailTemplater-win-x64.zip` (tek .exe) ve `MailTemplater.msix` üretir
(Actions > ilgili çalıştırma > Artifacts). `v1.0.0` gibi etiket atarsanız Release olarak da yayınlanır.

Store için Partner Center'dan uygulama adı rezerve edin, "Product identity" sayfasındaki
Package/Identity/Name, Package/Identity/Publisher ve PublisherDisplayName değerlerini
repo Settings > Secrets and variables > Actions > Variables altına
`IDENTITY_NAME`, `PUBLISHER`, `PUBLISHER_DISPLAY` olarak ekleyin. Store paketi sizin yerinize imzalar.
Not: İmzasız .msix doğrudan çift tıkla kurulmaz; yerel test için .zip içindeki .exe'yi kullanın.

## Logo
`Assets/logo.png` ve `Assets/app.ico` (zarf + şablon yığını + artı rozeti). Değiştirmek için aynı adlarla üzerine yazın;
Store görselleri `Assets/Square*.png`, `Wide310x150Logo.png`, `StoreLogo.png` dosyalarıdır.

## GitHub Actions ile derleme
1. Depoyu GitHub'a gönderin (`git push`). `.github/workflows/build.yml` otomatik çalışır.
2. Actions sekmesi > son çalışma > **Artifacts**: `MailTemplater-exe` ve `MailTemplater-msix`.
3. Release için: `git tag v1.0.0 && git push origin v1.0.0` (dosyalar Releases'e eklenir).
4. Store için `Packaging/AppxManifest.xml` içindeki 3 `REPLACE_...` değerini Partner Center'daki
   Product identity bilgileriyle değiştirin. Paket imzasızdır; Store yükleme sırasında imzalar.
   Yerelde denemek için ayrıca sertifika ile imzalamak gerekir.
5. Logo: `Assets/logo-1024.png` kaynak görseldir; değiştirmek için `app.ico` ve `Square*/StoreLogo/Wide*` PNG'lerini yeniden üretin.
