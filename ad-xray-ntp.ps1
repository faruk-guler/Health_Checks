#Requires -Version 2
<#
.SYNOPSIS
Active Directory Domain Controller'ların NTP (Saat Senkronizasyonu) Durumunu HTML olarak raporlar.

.DESCRIPTION
Active Directory Forest'taki her Domain Controller için W32TM komutlarını çalıştırır (Durum ve Yapılandırma) ve sonuçları tek bir HTML raporunda sunar.
#>

#---------------------------------------------------------[Değişkenler]--------------------------------------------------------

$ReportFolder = "C:\ADxRay"
if ((Test-Path -Path $ReportFolder -PathType Container) -eq $false) {
    New-Item -Type Directory -Force -Path $ReportFolder | Out-Null
}

$Global:ReportPath = "$ReportFolder\NTP_Report_" + (Get-Date -Format 'yyyy-MM-dd-hh-mm') + ".htm"
$Global:HTMLContent = New-Object System.Collections.Generic.List[string]

# Hata İşleme
$ErrorActionPreference = "SilentlyContinue"

#---------------------------------------------------------[Fonksiyonlar]--------------------------------------------------------

function Get-DCList {
    # Hata oluşursa null döner
    try {
        # Domain Controller listesini alır
        # ActiveDirectory modülünün yüklü olmasını gerektirir
        $DCs = Get-ADDomainController -Filter * | Select-Object Name, Domain
        return $DCs
    }
    catch {
        Write-Host "Hata: Active Directory'den DC listesi alınamadı. (AD Modülü yüklü olmayabilir veya yetki sorunu olabilir.)" -ForegroundColor Red
        # DC listesi alınamazsa, yerel makineyi dener (bu genelde DC üzerinde çalıştırılıyorsa)
        $LocalDC = Get-ADDomainController -Server (Get-ADDomain).PDCRoleOwner | Where-Object {$_.Name -eq $env:COMPUTERNAME}
        if ($LocalDC) {
            Write-Host "Yerel makineyi ('$($LocalDC.Name)') kontrol etmeye devam ediliyor." -ForegroundColor Yellow
            return @($LocalDC | Select-Object Name, Domain)
        }
        return $null
    }
}

function Get-NTPReport {
    param(
        [Parameter(Mandatory=$true)]
        [String]$DCName
    )
    
    Write-Host "`n--- $DCName için NTP Raporu Alınıyor... ---" -ForegroundColor Yellow

    # Sonuçları depolamak için dizi
    $ReportData = @()
    $StatusClass = "info"
    $ErrorMessage = $null

    try {
        # 1. NTP Durumu (Status)
        $StatusResult = Invoke-Command -ComputerName $DCName -ScriptBlock {
            w32tm /query /status | Out-String
        }
        
        # 2. NTP Yapılandırması (Configuration)
        $ConfigResult = Invoke-Command -ComputerName $DCName -ScriptBlock {
            w32tm /query /configuration | Out-String
        }

        # Olası önemli durumları kontrol et (Bu kısmı isteğe göre geliştirilebilir)
        if ($StatusResult -match "Source:\s+Local CMOS Clock") {
            $StatusClass = "warning"
        }
        if ($StatusResult -match "Last Sync Error:\s+0x[0-9a-fA-F]{8}") {
             if ($matches[0] -ne "Last Sync Error: 0x0") {
                $StatusClass = "error"
             }
        }
        
        $ReportData += $StatusResult
        $ReportData += $ConfigResult

    }
    catch {
        $ErrorMessage = "HATA: $DCName'e ulaşılamadı veya komut çalıştırılamadı: $($_.Exception.Message)"
        $StatusClass = "error"
        $ReportData += $ErrorMessage
    }
    
    # HTML Blokunu Oluştur
    $HTMLBlock = @"
<div class='card $StatusClass'>
    <h2>$DCName - NTP / Saat Senkronizasyonu Raporu</h2>
    <div class='tag $StatusClass'>Durum: $(if($StatusClass -eq "error"){"Hata/Ulaşılamaz"}elseif($StatusClass -eq "warning"){"Uyarı/Local CMOS"}else{"Başarılı"})</div>
    <div class='content'>
        <pre>
$(if ($ErrorMessage) {$ErrorMessage} else {$ReportData | Out-String})
        </pre>
    </div>
</div>
"@
    # Global HTML listesine ekle
    $Global:HTMLContent.Add($HTMLBlock)

    # Konsola kısa bilgi verme
    if ($StatusClass -eq "error") {
        Write-Host "Hata: $DCName raporu HTML'e eklendi. (Ulaşılamıyor)" -ForegroundColor Red
    } elseif ($StatusClass -eq "warning") {
        Write-Host "Uyarı: $DCName raporu HTML'e eklendi. (Local CMOS saati kullanılıyor olabilir)" -ForegroundColor DarkYellow
    } else {
        Write-Host "Başarılı: $DCName raporu HTML'e eklendi." -ForegroundColor Green
    }
}

function Generate-HTMLReport {
    
    # CSS Stil Tanımlamaları
    $CSS = @"
<style>
body { font-family: Arial, sans-serif; background-color: #f4f7f6; color: #333; margin: 0; padding: 20px; }
.header { text-align: center; background-color: #004c3d; color: white; padding: 20px; margin-bottom: 20px; border-radius: 8px; }
h1 { margin: 0; }
.container { max-width: 1000px; margin: auto; }
.card { 
    border: 1px solid #ddd; 
    border-radius: 8px; 
    margin-bottom: 20px; 
    box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    background-color: #fff;
    border-left: 8px solid;
}
.card h2 { 
    margin: 0; 
    padding: 15px; 
    font-size: 1.2em; 
    background-color: #f9f9f9;
    border-top-right-radius: 6px;
    border-top-left-radius: 6px;
}
.card.info { border-left-color: #007644; } /* Success/Info Color */
.card.warning { border-left-color: #ffc107; } /* Warning Color */
.card.error { border-left-color: #dc3545; } /* Error Color */

.content { padding: 15px; }
pre { 
    background-color: #eee; 
    padding: 15px; 
    border-radius: 5px; 
    overflow-x: auto; 
    white-space: pre-wrap;
    font-size: 0.9em;
    border: 1px solid #ccc;
}
.tag {
    display: inline-block;
    padding: 4px 10px;
    margin-left: 15px;
    margin-top: -10px;
    border-radius: 4px;
    font-weight: bold;
    font-size: 0.8em;
    color: white;
}
.tag.info { background-color: #007644; }
.tag.warning { background-color: #ffc107; color: #333; }
.tag.error { background-color: #dc3545; }
</style>
"@

    # HTML Başlığı
    $HTMLHead = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Active Directory NTP Raporu</title>
    $CSS
</head>
<body>
    <div class='header'>
        <h1>Active Directory NTP / Saat Senkronizasyonu Raporu</h1>
        <p>Rapor Tarihi: $(Get-Date)</p>
    </div>
    <div class='container'>
"@
    
    # HTML Kapanışı
    $HTMLFoot = @"
    </div>
    <div style='text-align: center; margin-top: 30px; color: #666; font-size: 0.8em;'>
        --- Rapor Sonu ---
    </div>
</body>
</html>
"@
    
    # Tüm parçaları birleştir ve dosyaya yaz
    $HTMLHead | Out-File -FilePath $Global:ReportPath -Encoding UTF8
    $Global:HTMLContent | Out-File -FilePath $Global:ReportPath -Encoding UTF8 -Append
    $HTMLFoot | Out-File -FilePath $Global:ReportPath -Encoding UTF8 -Append

}

#---------------------------------------------------------[Ana Akış]--------------------------------------------------------

Write-Host "Active Directory NTP Raporlama Aracı Başlatılıyor (HTML Çıktı)..." -ForegroundColor Green

$DCs = Get-DCList

if ($DCs -ne $null) {
    
    foreach ($DC in $DCs) {
        Get-NTPReport -DCName $DC.Name
    }
    
    # Raporlama tamamlandıktan sonra HTML dosyasını oluştur
    Generate-HTMLReport
    
} else {
    $DCs = @() # Liste boşsa, aşağıdaki bilgi mesajını engellemek için boş dizi yap
}

# Bitiş ve Rapor Yolu Bilgisi
if ($DCs.Count -gt 0 -or $Global:HTMLContent.Count -gt 0) {
    Write-Host "`nİşlem Tamamlandı." -ForegroundColor Green
    Write-Host "`n[INFO] Detaylı HTML raporu şuraya kaydedildi: $Global:ReportPath" -ForegroundColor Cyan
    Write-Host "Bu dosyayı tarayıcınızda açarak raporu görüntüleyebilirsiniz." -ForegroundColor Cyan
} else {
    Write-Host "`n[HATA] Hiçbir Domain Controller bulunamadı veya ulaşılamadı. Rapor oluşturulamadı." -ForegroundColor Red
}
