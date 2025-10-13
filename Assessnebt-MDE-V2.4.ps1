<#
.SYNOPSIS
    🛡️ MDE Security Assessment Tool v2.4 - Complete Microsoft Defender for Endpoint compliance assessment

.DESCRIPTION
    Professional security audit script that evaluates endpoint configurations against Microsoft best practices,
    Intune baseline, and advanced Windows controls, with enhanced visual interface and detailed reports.

    ╔══════════════════════════════════════════════════════════════════════════════════╗
    ║                        ASSESSED CATEGORIES (46+ controls)                        ║
    ╚══════════════════════════════════════════════════════════════════════════════════╝

    🛡️ MDE NATIVE - DEFENDER FOR ENDPOINT NATIVE IMPLEMENTATIONS:
    ────────────────────────────────────────────────────────────────────────────────────
    • MDE/EDR Onboarding (Sense service)
    • Microsoft Defender Antivirus:
      - Real-time protection, cloud-delivered protection, MAPS
      - PUA (Potentially Unwanted Applications)
      - Network Protection, Behavior Monitoring
      - Block at First Sight (BAFS)
      - Tamper Protection
    • Attack Surface Reduction (ASR) - 19 official rules with GUID and action code
    • Scan configurations (scheduling, parameters)
    • Signature updates and performance

    🪟 WINDOWS - SYSTEM SETTINGS (non-MDE specific):
    ────────────────────────────────────────────────────────────────────────────────────
    • Encryption: BitLocker (OS and removable media)
    • Windows Firewall (profiles and auditing)
    • Windows Defender SmartScreen (system-level protection - OS level)
    • Credential Guard (credential protection)
    • Secure Boot and TPM 2.0
    • WDAC (Windows Defender Application Control)
    • Exploit Protection (exploit mitigation)
    • Application Guard (browser isolation)
    • UAC (User Account Control)
    • LAPS (Local Administrator Password Solution)
    • Account policies (password, lockout)
    • Auditing and logs
    • Windows Update

    📋 COMPLIANCE - DEVICE & HARDWARE COMPLIANCE:
    ────────────────────────────────────────────────────────────────────────────────────
    • Device policies: minimum password length, screen lock
    • Windows Hello for Business (biometric authentication)
    • Updated TPM driver

    🌐 BROWSERS - BROWSER PROTECTION:
    ────────────────────────────────────────────────────────────────────────────────────
    • Microsoft Edge SmartScreen
    • Google Chrome Safe Browsing
    • Mozilla Firefox Enhanced Tracking Protection

    ╔══════════════════════════════════════════════════════════════════════════════════╗
    ║                            VERSION 2.4 FEATURES                                  ║
    ╚══════════════════════════════════════════════════════════════════════════════════╝

    ✨ NEW IN v2.4:
    • Complete reorganization: clear separation between MDE Native and Windows
    • Firefox protection (Enhanced Tracking Protection)
    • Professional visual interface with boxes, icons, and progress bars
    • Compliance score with graphical visualization by category
    • Real-time progress indicators during execution
    • Smart grouping of non-compliant items by category
    • Validated and updated documentation links (34 URLs)
    • Enhanced export with visual feedback (XLSX or CSV)
    • Error handling with solution suggestions

    📊 OUTPUT AND REPORTS:
    • Console organized in thematic sections with visual boxes
    • Overall and category compliance scores (with progress bars)
    • XLSX export (2 tabs: Complete checklist + Recommendations)
    • CSV fallback if ImportExcel module not available
    • All recommendations include: official Microsoft documentation, remediation command, and supplementary information

.PARAMETER OutputPath
    Optional full path for the report file.
    If omitted, automatically generates to C:\temp\MDE_Assessment_Report_[timestamp].xlsx

    Example: -OutputPath "C:\Reports\MDE_Assessment.xlsx"

.EXAMPLE
    .\Assessment-MDE-V2.4-EN.ps1

    Runs complete assessment with default output to C:\temp

.EXAMPLE
    .\Assessment-MDE-V2.4-EN.ps1 -OutputPath "C:\Reports\Assessment.xlsx"

    Runs with custom path for the report

.EXAMPLE
    .\Assessment-MDE-V2.4-EN.ps1 -Verbose

    Runs with progress details in console (verbose mode)

.NOTES
    ╔══════════════════════════════════════════════════════════════════════════════════╗
    ║                            SCRIPT INFORMATION                                    ║
    ╚══════════════════════════════════════════════════════════════════════════════════╝

    Version:         2.4
    Date:            10/07/2025
    Author:          Leandro Sardim
    Last Update:     MDE/Windows reorganization + Firefox + Professional Visual Interface

    REQUIREMENTS:
    ✓ PowerShell 5.1 or higher
    ✓ Run as Administrator (recommended for full access)
    ✓ Windows 10/11 with Microsoft Defender enabled
    ✓ Internet connection (for automatic ImportExcel module installation, if needed)
    ✓ Execution policy: automatically adjusted if needed

    VERSION HISTORY:
    • v2.4 (10/07/2025): Category reorganization + Firefox + Professional visual interface
    • v2.2 (previous): Control expansion + documentation + column separation
    • v2.0 (previous): Complete baseline with 40+ controls
    • v1.0 (previous): Initial version

    USEFUL LINKS:
    • MDE Documentation: https://learn.microsoft.com/en-us/defender-endpoint/
    • ASR Rules: https://learn.microsoft.com/en-us/defender-endpoint/attack-surface-reduction
    • Intune Baselines: https://learn.microsoft.com/en-us/mem/intune/protect/security-baselines

.LINK
    https://learn.microsoft.com/en-us/defender-endpoint/

.LINK
    https://learn.microsoft.com/en-us/windows/security/
#>

[CmdletBinding()]
param (
    [string]$OutputPath = ''
)

# Generate unique filename with date/time
$now = Get-Date -Format 'yyyy-MM-dd_HH-mm'
$reportFolder = 'C:\temp'
if (-not (Test-Path -Path $reportFolder)) {
    New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
}
$OutputPath = Join-Path $reportFolder "MDE_Assessment_Report_${now}.csv"

$ErrorActionPreference = 'Stop'

Write-Verbose 'Starting Microsoft Defender for Endpoint assessment.'

function Add-AssessmentResult {
    param (
        [ref]$Collector,
        [string]$Category,
        [string]$Setting,
        [string]$CurrentValue,
        [string]$BestPractice,
        [string]$Recommendation,
        [string]$Remediation,
        [bool]$Compliant,
        [string]$Documentation = $null,
        [string]$Guid = $null,
        [Nullable[int]]$CurrentCode = $null,
        [Nullable[int]]$RecommendationCode = $null
    )

    $Collector.Value += [PSCustomObject]@{
        Category = $Category
        'Configuration' = $Setting
        'Guid' = $Guid
        'Current Value' = $CurrentValue
        'Current Code' = $CurrentCode
        'Recommended Code' = $RecommendationCode
        'Best Practice' = $BestPractice
        'Recommendation' = $Recommendation
        'Remediation Command' = $(if ($Remediation -match '(.*)\. Guide: (.*)') { $matches[1].Trim() } else { $Remediation })
        'Documentation' = $Documentation
        'Supplementary Info' = $(if ($Remediation -match '(.*)\. Guide: (.*)') { $matches[2].Trim() } else { $null })
        Status = $(if ($Compliant) { 'Compliant' } else { 'Attention' })
    }
}

# Initial banner
Clear-Host
Write-Host
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
$title1 = '🛡️  MDE SECURITY ASSESSMENT TOOL v2.4'
$padLength1 = ($bannerWidth - $title1.Length) / 2
$padLeft1 = ' ' * [math]::Floor($padLength1)
$padRight1 = ' ' * [math]::Ceiling($padLength1)
Write-Host "║$padLeft1$title1$padRight1║" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
$subtitle1 = 'Microsoft Defender for Endpoint - Compliance Assessment'
$padLengthSub1 = ($bannerWidth - $subtitle1.Length) / 2
$padLeftSub1 = ' ' * [math]::Floor($padLengthSub1)
$padRightSub1 = ' ' * [math]::Ceiling($padLengthSub1)
Write-Host "║$padLeftSub1$subtitle1$padRightSub1║" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
Write-Host '  📋 Starting security assessment...' -ForegroundColor Yellow
Write-Host

Write-Verbose 'Loading current Defender preferences (Get-MpPreference, Get-MpComputerStatus)...'

try {
    Write-Host '  🔍 Collecting Microsoft Defender configurations...' -ForegroundColor Gray
    $mpPreference = Get-MpPreference
    $mpStatus = Get-MpComputerStatus
    Write-Verbose 'Preferences loaded successfully.'
    Write-Host '  ✓  Configurations collected successfully!' -ForegroundColor Green
    Write-Host
} catch {
    Write-Host
    Write-Host '╔══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Red
    Write-Host '║                              ❌ CRITICAL ERROR                                   ║' -ForegroundColor Red
    Write-Host '╚══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Red
    Write-Host
    Write-Host '  ⚠️  Unable to obtain Microsoft Defender configurations.' -ForegroundColor Yellow
    Write-Host '  💡 Solution: Run the script as administrator.' -ForegroundColor Cyan
    Write-Host
    exit 1
}

$tamperProtectionKey = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features'
$tamperProtection = Get-ItemProperty -Path $tamperProtectionKey -Name TamperProtection -ErrorAction SilentlyContinue


$assessmentResults = @()

# Microsoft Intune baseline (recommended values)
$intuneBaseline = @{
    'Defender_RealTimeProtection' = $true
    'Defender_CloudProtection' = 2
    'Defender_PUAProtection' = 1
    'Defender_NetworkProtection' = 1
    'Defender_BehaviorMonitoring' = $true
    'Defender_TamperProtection' = 5
    'Defender_ScanAvgCPULoadFactor' = 50
    'Defender_ScanScheduleDay' = 0
    'Defender_ScanScheduleTime' = 120
    'BitLocker_Protection' = $true
    'Firewall_Enabled' = $true
}

# ======= MDE NATIVE IMPLEMENTATIONS =======

# 1. MDE/EDR onboarding status (Sense service)
$mdeService = Get-Service -Name Sense -ErrorAction SilentlyContinue
$mdeOnboarded = $mdeService -and $mdeService.Status -eq 'Running'
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Native' -Setting 'MDE/EDR Onboarding' `
    -CurrentValue $(if ($mdeOnboarded) { 'Onboarded (Sense running)' } else { 'Not onboarded' }) `
    -BestPractice 'Onboarded (Sense service active)' `
    -Recommendation 'Ensure device is onboarded to Microsoft Defender for Endpoint.' `
    -Documentation 'https://learn.microsoft.com/microsoft-365/security/defender-endpoint/onboarding' `
    -Remediation 'Run onboarding script or apply Intune MDE policy' `
    -Compliant $mdeOnboarded

# ======= END MDE NATIVE (more MDE native configurations will be grouped after) =======

# ======= WINDOWS SETTINGS (non-MDE specific) =======
# Minimum password length
$passwordPolicy = (Get-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name MinimumPasswordLength -ErrorAction SilentlyContinue).MinimumPasswordLength
$passwordCompliant = $passwordPolicy -ge 8
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'Minimum password length' `
    -CurrentValue $(if ($passwordPolicy) { "$passwordPolicy characters" } else { 'Not configured' }) `
    -BestPractice '≥ 8 characters' `
    -Recommendation 'Set minimum password length to at least 8 characters (recommended: 12+).' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/security-policy-settings/minimum-password-length' `
    -Remediation 'Set-LocalUser -Name <user> -Password <strong password>; or set via Intune/GPO' `
    -Compliant $passwordCompliant

# Automatic screen lock (Screen Saver timeout)
$lockScreenTimeout = (Get-ItemProperty -Path "HKCU:Control Panel\Desktop" -Name ScreenSaveTimeOut -ErrorAction SilentlyContinue).ScreenSaveTimeOut
$lockScreenCompliant = $lockScreenTimeout -and ([int]$lockScreenTimeout -le 900)
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'Automatic screen lock' `
    -CurrentValue $(if ($lockScreenTimeout) { "$lockScreenTimeout seconds" } else { 'Not configured' }) `
    -BestPractice '≤ 900 seconds (15 min)' `
    -Recommendation 'Set automatic lock to 15 minutes or less of inactivity.' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/security-policy-settings/interactive-logon-machine-inactivity-limit' `
    -Remediation 'Set-ItemProperty -Path "HKCU:Control Panel\Desktop" -Name ScreenSaveTimeOut -Value 900; or set via Intune/GPO' `
    -Compliant $lockScreenCompliant

# Windows Hello (simple example – presence of FaceLogon key)
$helloKey = 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\FaceLogon'
$helloEnabled = (Get-ItemProperty -Path $helloKey -Name Enabled -ErrorAction SilentlyContinue).Enabled -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'Windows Hello (biometrics)' `
    -CurrentValue $(if ($helloEnabled) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Enable Windows Hello (PIN + biometrics) to strengthen authentication.' `
    -Documentation 'https://learn.microsoft.com/windows/security/identity-protection/hello-for-business/' `
    -Remediation 'Configure via Intune: Windows Hello for Business' `
    -SupplementaryInfo 'https://learn.microsoft.com/windows/security/identity-protection/hello-for-business/deploy/hybrid-cloud-kerberos-trust' `
    -Compliant $helloEnabled

# 3. Updated TPM driver (≤ 1 year)
$tpmDriver = Get-WmiObject -Class Win32_PnPSignedDriver -ErrorAction SilentlyContinue | Where-Object { $_.DeviceName -like '*TPM*' }
$tpmDriverDate = if ($tpmDriver) { $tpmDriver.DriverDate } else { $null }
$tpmDriverCompliant = $tpmDriverDate -and ((Get-Date) - $tpmDriverDate).Days -le 365
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'Updated TPM driver' `
    -CurrentValue $(if ($tpmDriverDate) { $tpmDriverDate.ToString('yyyy-MM-dd') } else { 'Not found' }) `
    -BestPractice 'Driver updated (≤ 1 year)' `
    -Recommendation 'Update TPM driver to ensure compatibility and fixes.' `
    -Documentation 'https://learn.microsoft.com/windows/security/hardware-security/tpm/trusted-platform-module-top-node' `
    -Remediation 'Update via Windows Update or manufacturer' `
    -Compliant $tpmDriverCompliant

# 7. LAPS (Local Admin Password Solution)
$lapsKey = 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS'
$lapsEnabled = (Get-ItemProperty -Path $lapsKey -Name Enable -ErrorAction SilentlyContinue).Enable -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'LAPS enabled' `
    -CurrentValue $(if ($lapsEnabled) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Enable Windows LAPS to protect local administrator credentials.' `
    -Documentation 'https://learn.microsoft.com/windows-server/identity/laps/laps-overview' `
    -Remediation 'Configure Windows LAPS policy via Intune / GPO' `
    -SupplementaryInfo 'https://learn.microsoft.com/windows-server/identity/laps/laps-scenarios-legacy' `
    -Compliant $lapsEnabled

# 8. SmartScreen Protection (System + Browsers)
# Windows Defender SmartScreen (system-level protection for apps and files)
$systemSmartScreen = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer' -Name SmartScreenEnabled -ErrorAction SilentlyContinue).SmartScreenEnabled
$systemSmartScreenCompliant = $systemSmartScreen -in @('Warn', 'Block')
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Windows Defender SmartScreen (System)' `
    -CurrentValue $(if ($systemSmartScreen) { $systemSmartScreen } else { 'Not configured' }) `
    -BestPractice 'Warn or Block' `
    -Recommendation 'Enable Windows SmartScreen to protect against malicious apps/files from any source.' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/microsoft-defender-smartscreen/microsoft-defender-smartscreen-overview' `
    -Remediation 'Set-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name SmartScreenEnabled -Value "Block"; or Intune/GPO policy' `
    -Compliant $systemSmartScreenCompliant

# Edge SmartScreen (browser-specific protection)
$edgeSmartScreen = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Microsoft\Edge' -Name SmartScreenEnabled -ErrorAction SilentlyContinue).SmartScreenEnabled -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Browsers' -Setting 'Microsoft Edge SmartScreen' `
    -CurrentValue $(if ($edgeSmartScreen) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Ensure SmartScreen is enabled in Edge to block malicious sites/downloads.' `
    -Documentation 'https://learn.microsoft.com/deployedge/microsoft-edge-security-smartscreen' `
    -Remediation 'Set-ItemProperty -Path "HKLM:SOFTWARE\Policies\Microsoft\Edge" -Name SmartScreenEnabled -Value 1; or Intune/GPO policy' `
    -Compliant $edgeSmartScreen

# Chrome Safe Browsing (basic level = 1; enhanced may be 2 depending on the organization)
$chromeSafeBrowsing = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Google\Chrome' -Name SafeBrowsingProtectionLevel -ErrorAction SilentlyContinue).SafeBrowsingProtectionLevel
$chromeSafeBrowsingEnabled = $chromeSafeBrowsing -in 1,2
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Browsers' -Setting 'Chrome Safe Browsing' `
    -CurrentValue $(if ($chromeSafeBrowsing) { "Level $chromeSafeBrowsing" } else { 'Not configured' }) `
    -BestPractice 'Enabled (level 1 or 2)' `
    -Recommendation 'Keep Safe Browsing at least at default level; consider enhanced.' `
    -Documentation 'https://support.google.com/chrome/answer/99020' `
    -Remediation 'Set-ItemProperty -Path "HKLM:SOFTWARE\Policies\Google\Chrome" -Name SafeBrowsingProtectionLevel -Value 1; or Intune/GPO policy' `
    -Compliant $chromeSafeBrowsingEnabled

# Firefox Enhanced Tracking Protection + Safe Browsing
# Firefox uses preferences.js, but corporate policies may be via HKLM\SOFTWARE\Policies\Mozilla\Firefox
$firefoxETP = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Mozilla\Firefox' -Name 'EnableTrackingProtection' -ErrorAction SilentlyContinue).EnableTrackingProtection -eq 1
$firefoxSafeBrowsing = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Mozilla\Firefox' -Name 'DisableSafeMode' -ErrorAction SilentlyContinue).DisableSafeMode -ne 1 # Safe Browsing enabled by default if not disabled
$firefoxProtectionEnabled = $firefoxETP -or $firefoxSafeBrowsing
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Browsers' -Setting 'Firefox Enhanced Tracking Protection' `
    -CurrentValue $(if ($firefoxETP) { 'Enabled' } elseif ($firefoxSafeBrowsing) { 'Safe Browsing active' } else { 'Not configured/Disabled' }) `
    -BestPractice 'Enabled (Enhanced Tracking Protection + Safe Browsing)' `
    -Recommendation 'Enable Enhanced Tracking Protection and Safe Browsing in Firefox for protection against tracking and malicious sites.' `
    -Documentation 'https://support.mozilla.org/kb/enhanced-tracking-protection-firefox-desktop' `
    -Remediation 'Configure via Mozilla corporate policy or about:config: privacy.trackingprotection.enabled=true' `
    -SupplementaryInfo 'https://support.mozilla.org/kb/how-does-phishing-and-malware-protection-work' `
    -Compliant $firefoxProtectionEnabled

# ======= END IMPROVEMENTS =======



# 1. Real-time protection (Baseline Intune)
$rtpEnabled = -not $mpPreference.DisableRealtimeMonitoring
$rtpCompliant = $rtpEnabled -eq $intuneBaseline['Defender_RealTimeProtection']
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Native' -Setting 'Real-time protection (Baseline Intune)' `
    -CurrentValue $(if ($rtpEnabled) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled (DisableRealtimeMonitoring = 0)' `
    -Recommendation 'Keep real-time protection always enabled.' `
    -Remediation 'Set-MpPreference -DisableRealtimeMonitoring $false' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/microsoft-defender-antivirus/configure-real-time-protection-microsoft-defender-antivirus' `
    -Compliant $rtpCompliant `
    -CurrentCode ([int]$mpPreference.DisableRealtimeMonitoring) `
    -RecommendationCode 0

# 21. BitLocker Protection (Baseline Intune)
$bitlockerStatus = $null
try {
    $bitlockerStatus = (Get-BitLockerVolume -MountPoint 'C:').ProtectionStatus
} catch {}
$bitlockerEnabled = $bitlockerStatus -eq 1
$bitlockerCompliant = $bitlockerEnabled -eq $intuneBaseline['BitLocker_Protection']
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'BitLocker Protected (Baseline Intune)' `
    -CurrentValue $(if ($bitlockerEnabled) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Enable BitLocker to protect data at rest.' `
    -Remediation 'Enable-BitLocker -MountPoint C:' `
    -Documentation 'https://learn.microsoft.com/windows/security/information-protection/bitlocker/bitlocker-overview' `
    -Compliant $bitlockerCompliant

# 22. Firewall Enabled (Baseline Intune)
$firewallEnabled = $null
try {
    $firewallEnabled = (Get-NetFirewallProfile | Where-Object { $_.Enabled -eq 'True' }).Count -ge 1
} catch {}
$firewallCompliant = $firewallEnabled -eq $intuneBaseline['Firewall_Enabled']
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Firewall Enabled (Baseline Intune)' `
    -CurrentValue $(if ($firewallEnabled) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Keep Windows Firewall enabled on all profiles.' `
    -Remediation 'Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True' `
    -Documentation 'https://learn.microsoft.com/windows/security/operating-system-security/network-security/windows-firewall/windows-firewall-with-advanced-security' `
    -Compliant $firewallCompliant

# 2. Cloud-delivered protection (MAPS)
$mapsLevel = $mpPreference.MAPSReporting
$mapsLabel = switch ($mapsLevel) {
    0 { 'Disabled' }
    1 { 'Basic' }
    2 { 'Advanced' }
    default { "Unknown ($mapsLevel)" }
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Cloud-delivered protection' `
    -CurrentValue $mapsLabel `
    -BestPractice 'Advanced (MAPSReporting = 2)' `
    -Recommendation 'Enable advanced level for real-time cloud response.' `
    -Remediation 'Set-MpPreference -MAPSReporting Advanced' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant ($mapsLevel -eq 2) `
    -CurrentCode $mapsLevel `
    -RecommendationCode 2

# 3. Automatic sample submission
$sampleConsent = $mpPreference.SubmitSamplesConsent
$sampleLabel = switch ($sampleConsent) {
    0 { 'Always ask' }
    1 { 'Automatically send safe samples' }
    2 { 'Never send' }
    3 { 'Automatically send all samples' }
    default { "Unknown ($sampleConsent)" }
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Automatic sample submission' `
    -CurrentValue $sampleLabel `
    -BestPractice 'Automatically send all samples (value 3)' `
    -Recommendation 'Ensure automatic sample submission for higher detection effectiveness.' `
    -Remediation 'Set-MpPreference -SubmitSamplesConsent SendAllSamples' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant ($sampleConsent -eq 3) `
    -CurrentCode $sampleConsent `
    -RecommendationCode 3

# 4. PUA Protection
$puaState = $mpPreference.PUAProtection
$puaLabel = switch ($puaState) {
    0 { 'Disabled' }
    1 { 'Enabled' }
    2 { 'Audit only' }
    default { "Unknown ($puaState)" }
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Potentially Unwanted Application (PUA) Protection' `
    -CurrentValue $puaLabel `
    -BestPractice 'Enabled' `
    -Recommendation 'Block potentially unwanted applications to reduce attack surface.' `
    -Remediation 'Set-MpPreference -PUAProtection Enabled' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/detect-block-potentially-unwanted-apps-microsoft-defender-antivirus' `
    -Compliant ($puaState -eq 1) `
    -CurrentCode $puaState `
    -RecommendationCode 1

# 5. Network protection
$networkProtection = $mpPreference.EnableNetworkProtection
$networkLabel = switch ($networkProtection) {
    0 { 'Disabled' }
    1 { 'Block' }
    2 { 'Audit only' }
    default { "Unknown ($networkProtection)" }
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Network protection' `
    -CurrentValue $networkLabel `
    -BestPractice 'Block (EnableNetworkProtection = 1)' `
    -Recommendation 'Set network protection to block mode to prevent malicious communications.' `
    -Remediation 'Set-MpPreference -EnableNetworkProtection Enabled' `
    -Documentation 'https://learn.microsoft.com/microsoft-365/security/defender-endpoint/network-protection' `
    -Compliant ($networkProtection -eq 1) `
    -CurrentCode $networkProtection `
    -RecommendationCode 1

# 6. Behavior monitoring
$behaviorMonitoring = -not $mpPreference.DisableBehaviorMonitoring
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Behavior monitoring' `
    -CurrentValue $(if ($behaviorMonitoring) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Detect suspicious behavior by analyzing real-time activities.' `
    -Remediation 'Set-MpPreference -DisableBehaviorMonitoring $false' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant $behaviorMonitoring `
    -CurrentCode ([int]$mpPreference.DisableBehaviorMonitoring) `
    -RecommendationCode 0

# 7. Block at First Sight
$blockAtFirstSeen = -not $mpPreference.DisableBlockAtFirstSeen
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Block at First Sight' `
    -CurrentValue $(if ($blockAtFirstSeen) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Allow Defender to block files never seen before as soon as they appear.' `
    -Remediation 'Set-MpPreference -DisableBlockAtFirstSeen $false' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/microsoft-defender-antivirus/configure-block-at-first-sight-microsoft-defender-antivirus' `
    -Compliant $blockAtFirstSeen `
    -CurrentCode ([int]$mpPreference.DisableBlockAtFirstSeen) `
    -RecommendationCode 0

# 8. Tamper Protection
$tamperValue = $tamperProtection.TamperProtection
$tamperEnabled = $tamperValue -eq 5
$tamperLabel = if ($null -ne $tamperValue) {
    if ($tamperEnabled) { 'Enabled (5)' } else { "Disabled ($tamperValue)" }
} else {
    'Not configured'
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Tamper Protection' `
    -CurrentValue $tamperLabel `
    -BestPractice 'Enabled (value 5)' `
    -Recommendation 'Enable via Microsoft Defender portal or Microsoft Intune.' `
    -Remediation 'Configuration via MDE/Intune portal (not available in PowerShell).' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/prevent-changes-to-security-settings-with-tamper-protection' `
    -Compliant $tamperEnabled `
    -Guid $(if ($null -ne $tamperValue) { 'TamperProtection' } else { $null }) `
    -CurrentCode $tamperValue `
    -RecommendationCode 5

# 9. Exclude network files from scanning
$networkScanDisabled = $mpPreference.DisableScanningNetworkFiles
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Ignore network files already protected' `
    -CurrentValue $(if ($networkScanDisabled) { 'Disabled (recommended)' } else { 'Enabled' }) `
    -BestPractice 'Disabled to avoid double scanning' `
    -Recommendation 'Disable scanning of network files to reduce performance impact (protection should exist on the source server).' `
    -Remediation 'Set-MpPreference -DisableScanningNetworkFiles $true' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant $networkScanDisabled `
    -CurrentCode ([int]$networkScanDisabled) `
    -RecommendationCode 1

# 10. Average CPU load during scans
$cpuLoad = $mpPreference.ScanAvgCPULoadFactor
$cpuDisplay = if ($cpuLoad -is [int]) { "$cpuLoad%" } else { 'Not configured' }
$cpuCompliant = ($cpuLoad -is [int]) -and ($cpuLoad -le 50)
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Average CPU load during scans' `
    -CurrentValue $cpuDisplay `
    -BestPractice '≤ 50% (recommended default value)' `
    -Recommendation 'Limit CPU load to ensure protection without degrading performance.' `
    -Remediation 'Set-MpPreference -ScanAvgCPULoadFactor 50' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/tune-performance-defender-antivirus' `
    -Compliant $cpuCompliant `
    -CurrentCode $(if ($cpuLoad -is [int]) { $cpuLoad } else { $null }) `
    -RecommendationCode 50

function Get-BoolLabel {
    param ($Value)

    if ($null -eq $Value) { return 'Not configured' }
    if ($Value -is [bool]) {
        return $(if ($Value) { 'Enabled' } else { 'Disabled' })
    }

    switch ($Value) {
        0 { 'Disabled (0)' }
        1 { 'Enabled (1)' }
        default { $Value.ToString() }
    }
}

function Get-CloudBlockLevelLabel {
    param ([Nullable[int]]$Value)

    if ($null -eq $Value) { return 'Not configured' }

    switch ($Value) {
        0 { 'Disabled (0)' }
        1 { 'Low (1)' }
        2 { 'High (2)' }
        3 { 'High Plus (3)' }
        4 { 'Zero Tolerance (4)' }
        default { "Unknown ($Value)" }
    }
}

function Get-UnknownThreatActionLabel {
    param ([Nullable[int]]$Value)

    if ($null -eq $Value) { return 'Not configured' }

    switch ($Value) {
        0 { 'Allow (0)' }
        1 { 'Block (1)' }
        2 { 'Quarantine (2)' }
        3 { 'Remove (3)' }
        4 { 'Clean (4)' }
        5 { 'Ignore (5)' }
        6 { 'Block (recommended parameter)' }
        default { "Unknown ($Value)" }
    }
}

function Get-ScanTypeLabel {
    param ([Nullable[int]]$Value)

    if ($null -eq $Value) { return 'Not configured' }

    switch ($Value) {
        1 { 'Quick (1)' }
        2 { 'Full (2)' }
        3 { 'Custom (3)' }
        default { "Unknown ($Value)" }
    }
}

function Get-ScanDayLabel {
    param ([Nullable[int]]$Value)

    if ($null -eq $Value) { return 'Every day (0)' }

    switch ($Value) {
        0 { 'Every day (0)' }
        1 { 'Sunday (1)' }
        2 { 'Monday (2)' }
        3 { 'Tuesday (3)' }
        4 { 'Wednesday (4)' }
        5 { 'Thursday (5)' }
        6 { 'Friday (6)' }
        7 { 'Saturday (7)' }
        default { "Unknown ($Value)" }
    }
}

function Convert-MinutesToTimeLabel {
    param ([Nullable[int]]$Minutes)

    if ($null -eq $Minutes) { return 'Not configured' }
    if ($Minutes -lt 0) { return "Unknown ($Minutes)" }

    $hours = [int]($Minutes / 60)
    $mins = $Minutes % 60
    '{0:00}:{1:00} ({2})' -f $hours, $mins, $Minutes
}

function Get-AsrActionLabel {
    param ([Nullable[int]]$Action)

    if ($null -eq $Action) { return 'Not configured' }

    switch ($Action) {
        0 { 'Disabled (0)' }
        1 { 'Block (1)' }
        2 { 'Audit only (2)' }
        3 { 'Disable (3)' }
        4 { 'Block - legacy (4)' }
        5 { 'Allow (5)' }
        6 { 'Notify (6)' }
        default { "Unknown ($Action)" }
    }
}

# 11. Allow Network Protection Down-Level
$allowNpd = $mpPreference.AllowNetworkProtectionDownLevel
$allowNpdValue = if ($null -eq $allowNpd) { $null } else { [int]$allowNpd }
$allowNpdEnabled = $allowNpdValue -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'AllowNetworkProtectionDownLevel' `
    -CurrentValue (Get-BoolLabel $allowNpd) `
    -BestPractice 'Enabled (value 1)' `
    -Recommendation 'Enable network protection also for down-level Windows versions.' `
    -Remediation 'Set-MpPreference -AllowNetworkProtectionDownLevel $true' `
    -Documentation 'https://learn.microsoft.com/microsoft-365/security/defender-endpoint/network-protection' `
    -Compliant $allowNpdEnabled `
    -CurrentCode $allowNpdValue `
    -RecommendationCode 1

# 12. Allow Network Protection on Windows Server
$allowNpw = $mpPreference.AllowNetworkProtectionOnWinServer
$allowNpwValue = if ($null -eq $allowNpw) { $null } else { [int]$allowNpw }
$allowNpwEnabled = $allowNpwValue -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'AllowNetworkProtectionOnWinServer' `
    -CurrentValue (Get-BoolLabel $allowNpw) `
    -BestPractice 'Enabled (value 1)' `
    -Recommendation 'Enable network protection on servers to expand blocking coverage.' `
    -Remediation 'Set-MpPreference -AllowNetworkProtectionOnWinServer $true' `
    -Documentation 'https://learn.microsoft.com/microsoft-365/security/defender-endpoint/network-protection' `
    -Compliant $allowNpwEnabled `
    -CurrentCode $allowNpwValue `
    -RecommendationCode 1

# 13. Cloud Block Level
$cloudBlockLevel = $mpPreference.CloudBlockLevel
$cloudBlockCompliant = $cloudBlockLevel -eq 2
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'CloudBlockLevel' `
    -CurrentValue (Get-CloudBlockLevelLabel $cloudBlockLevel) `
    -BestPractice 'High (value 2)' `
    -Recommendation 'Set cloud block level to High for more aggressive response.' `
    -Remediation "Set-MpPreference -CloudBlockLevel High" `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant $cloudBlockCompliant `
    -CurrentCode $cloudBlockLevel `
    -RecommendationCode 2

# 14. Cloud Extended Timeout
$cloudExtendedTimeout = $mpPreference.CloudExtendedTimeout
$cloudTimeoutCompliant = ($cloudExtendedTimeout -is [int]) -and ($cloudExtendedTimeout -ge 50)
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'CloudExtendedTimeout' `
    -CurrentValue $(if ($cloudExtendedTimeout -is [int]) { "$cloudExtendedTimeout seconds" } else { 'Not configured' }) `
    -BestPractice '50 seconds (recommended value)' `
    -Recommendation 'Allow cloud sufficient time to analyze files before releasing.' `
    -Remediation 'Set-MpPreference -CloudExtendedTimeout 50' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant $cloudTimeoutCompliant `
    -CurrentCode $(if ($cloudExtendedTimeout -is [int]) { $cloudExtendedTimeout } else { $null }) `
    -RecommendationCode 50

# 15. Unknown Threat Default Action
$unknownThreatAction = $mpPreference.UnknownThreatDefaultAction
$unknownThreatCompliant = $unknownThreatAction -eq 6
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'UnknownThreatDefaultAction' `
    -CurrentValue (Get-UnknownThreatActionLabel $unknownThreatAction) `
    -BestPractice 'Block/Quarantine (value 6)' `
    -Recommendation 'Ensure unknown threats are automatically blocked.' `
    -Remediation 'Set-MpPreference -UnknownThreatDefaultAction Block' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant $unknownThreatCompliant `
    -CurrentCode $unknownThreatAction `
    -RecommendationCode 6

# 16. Check for signatures before scan
$checkBeforeScan = $mpPreference.CheckForSignaturesBeforeRunningScan
$checkBeforeScanValue = if ($null -eq $checkBeforeScan) { $null } else { [int]$checkBeforeScan }
$checkBeforeScanCompliant = $checkBeforeScanValue -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'CheckForSignaturesBeforeRunningScan' `
    -CurrentValue (Get-BoolLabel $checkBeforeScan) `
    -BestPractice 'Enabled (value 1)' `
    -Recommendation 'Check for new signatures before starting any scheduled scan.' `
    -Remediation 'Set-MpPreference -CheckForSignaturesBeforeRunningScan $true' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/microsoft-defender-antivirus-updates' `
    -Compliant $checkBeforeScanCompliant `
    -CurrentCode $checkBeforeScanValue `
    -RecommendationCode 1

# 17. Scheduled scan parameters
$scanParameters = $mpPreference.ScanParameters
$scanParametersCompliant = $scanParameters -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'ScanParameters' `
    -CurrentValue (Get-ScanTypeLabel $scanParameters) `
    -BestPractice 'Quick daily scan (value 1)' `
    -Recommendation 'Use quick scan for automatic schedules.' `
    -Remediation 'Set-MpPreference -ScanParameters 1' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/schedule-antivirus-scans' `
    -Compliant $scanParametersCompliant `
    -CurrentCode $scanParameters `
    -RecommendationCode 1

# 18. Scheduled scan day
$scanScheduleDay = $mpPreference.ScanScheduleDay
$scanScheduleDayCompliant = $scanScheduleDay -eq 0
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'ScanScheduleDay' `
    -CurrentValue (Get-ScanDayLabel $scanScheduleDay) `
    -BestPractice 'Every day (value 0)' `
    -Recommendation 'Schedule daily scan for continuous coverage.' `
    -Remediation 'Set-MpPreference -ScanScheduleDay 0' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/schedule-antivirus-scans' `
    -Compliant $scanScheduleDayCompliant `
    -CurrentCode $scanScheduleDay `
    -RecommendationCode 0

# 19. Scheduled scan time
$scanScheduleTime = $mpPreference.ScanScheduleTime
$scanScheduleTimeMinutes = if ($scanScheduleTime -is [TimeSpan]) { 
    [int]($scanScheduleTime.TotalMinutes) 
} elseif ($scanScheduleTime -is [int]) { 
    $scanScheduleTime 
} else { 
    $null 
}
$scanScheduleTimeCompliant = ($scanScheduleTimeMinutes -eq 120)
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'ScanScheduleTime' `
    -CurrentValue (Convert-MinutesToTimeLabel $scanScheduleTimeMinutes) `
    -BestPractice '02:00 (value 120)' `
    -Recommendation 'Keep scheduled scan at low usage time.' `
    -Remediation 'Set-MpPreference -ScanScheduleTime 120' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/schedule-antivirus-scans' `
    -Compliant $scanScheduleTimeCompliant `
    -CurrentCode $scanScheduleTimeMinutes `
    -RecommendationCode 120

Write-Host '  ⚙️  Checking main Defender configurations...' -ForegroundColor Gray
Write-Verbose 'Main configuration checklist completed. Starting ASR rules assessment.'

# Map current ASR rule actions
$asrIds = @()
$asrActions = @()
if ($mpPreference.AttackSurfaceReductionRules_Ids) { $asrIds = $mpPreference.AttackSurfaceReductionRules_Ids }
if ($mpPreference.AttackSurfaceReductionRules_Actions) { $asrActions = $mpPreference.AttackSurfaceReductionRules_Actions }

$asrActionById = @{}
for ($i = 0; $i -lt $asrIds.Count; $i++) {
    $id = $asrIds[$i]
    if ($null -ne $id) {
        $asrActionById[$id.ToString().ToUpper()] = $asrActions[$i]
    }
}

$asrCatalog = @(
    @{ Id = '56a863a9-875e-4185-98a7-b882c64b5ce5'; Name = 'Block abuse of exploited vulnerable signed drivers'; Recommended = 1; Recommendation = 'Bloquear drivers assinados explorados para evitar elevação de privilégio.' }
    @{ Id = '7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c'; Name = 'Block Adobe Reader from creating child processes'; Recommended = 1; Recommendation = 'Impedir que o Adobe Reader seja usado como vetor para executar cargas maliciosas.' }
    @{ Id = 'd4f940ab-401b-4efc-aadc-ad5f3c50688a'; Name = 'Block Office applications from creating child processes'; Recommended = 1; Recommendation = 'Evitar que Office lance processos suspeitos através de macros.' }
    @{ Id = '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2'; Name = 'Block credential stealing from LSASS'; Recommended = 1; Recommendation = 'Impedir que processos acessem LSASS para roubo de credenciais.' }
    @{ Id = 'be9ba2d9-53ea-4cdc-84e5-9b1eeee46550'; Name = 'Block executable content from email and webmail'; Recommended = 1; Recommendation = 'Bloquear execução de anexos potencialmente maliciosos oriundos de e-mail.' }
    @{ Id = '01443614-cd74-433a-b99e-2ecdc07bfc25'; Name = 'Block executable files unless they meet prevalence/age/trusted list criteria'; Recommended = 1; Recommendation = 'Permitir apenas executáveis confiáveis com base em reputação de nuvem.' }
    @{ Id = '5beb7efe-fd9a-4556-801d-275e5ffc04cc'; Name = 'Block execution of potentially obfuscated scripts'; Recommended = 1; Recommendation = 'Bloquear scripts ofuscados que podem esconder código malicioso.' }
    @{ Id = 'd3e037e1-3eb8-44c8-a917-57927947596d'; Name = 'Block JavaScript or VBScript from launching downloaded executable content'; Recommended = 1; Recommendation = 'Evitar que scripts baixados lancem executáveis diretamente.' }
    @{ Id = '3b576869-a4ec-4529-8536-b80a7769e899'; Name = 'Block Office applications from creating executable content'; Recommended = 1; Recommendation = 'Impedir persistência por meio de conteúdo executável gerado pelo Office.' }
    @{ Id = '75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84'; Name = 'Block Office applications from injecting code into other processes'; Recommended = 1; Recommendation = 'Evitar que Office injete código em outros processos.' }
    @{ Id = '26190899-1602-49e8-8b27-eb1d0a1ce869'; Name = 'Block Office communication applications from creating child processes'; Recommended = 1; Recommendation = 'Bloquear que Outlook e apps de comunicação criem processos suspeitos.' }
    @{ Id = 'e6db77e5-3df2-4cf1-b95a-636979351e5b'; Name = 'Block persistence through WMI event subscription'; Recommended = 1; Recommendation = 'Prevenir que malwares abusem de inscrições WMI para persistência.' }
    @{ Id = 'd1e49aac-8f56-4280-b9ba-993a6d77406c'; Name = 'Block process creations originating from PSExec and WMI commands'; Recommended = 1; Recommendation = 'Restringir uso malicioso de PsExec ou WMI para executar código.' }
    @{ Id = '33ddedf1-c6e0-47cb-833e-de6133960387'; Name = 'Block rebooting machine in Safe Mode'; Recommended = 1; Recommendation = 'Evitar que ameaças reiniciem a máquina em modo seguro para burlar proteções.' }
    @{ Id = 'b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4'; Name = 'Block untrusted and unsigned processes that run from USB'; Recommended = 1; Recommendation = 'Bloquear execução de binários não confiáveis vindos de mídia removível.' }
    @{ Id = 'c0033c00-d16d-4114-a5a0-dc9b3a7d2ceb'; Name = 'Block use of copied or impersonated system tools'; Recommended = 1; Recommendation = 'Impedir uso de ferramentas do sistema falsificadas para evasão.' }
    @{ Id = 'a8f5898e-1dc8-49a9-9878-85004b8a61e6'; Name = 'Block Webshell creation for Servers'; Recommended = 1; Recommendation = 'Bloquear criação de webshells em servidores com funções web/Exchange.' }
    @{ Id = '92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b'; Name = 'Block Win32 API calls from Office macros'; Recommended = 1; Recommendation = 'Evitar que macros chamem APIs Win32 para executar payloads.' }
    @{ Id = 'c1db55ab-c21a-4637-bb3f-a12568109d35'; Name = 'Use advanced protection against ransomware'; Recommended = 1; Recommendation = 'Habilitar heurísticas avançadas contra ransomware.' }
)

foreach ($rule in $asrCatalog) {
    $ruleIdKey = $rule.Id.ToUpper()
    $currentAction = if ($asrActionById.ContainsKey($ruleIdKey)) { $asrActionById[$ruleIdKey] } else { $null }
    $currentLabel = Get-AsrActionLabel $currentAction
    $bestLabel = Get-AsrActionLabel $rule.Recommended
    $isCompliant = $currentAction -eq $rule.Recommended
    $remediation = "Set-MpPreference -AttackSurfaceReductionRules_Ids $($rule.Id) -AttackSurfaceReductionRules_Actions $($rule.Recommended)"

    Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting $rule.Name `
        -CurrentValue $currentLabel `
        -BestPractice $bestLabel `
        -Recommendation $rule.Recommendation `
        -Remediation $remediation `
        -Documentation 'https://learn.microsoft.com/microsoft-365/security/defender-endpoint/attack-surface-reduction' `
        -Compliant $isCompliant `
        -Guid $rule.Id `
        -CurrentCode $currentAction `
        -RecommendationCode $rule.Recommended
}

Write-Host '  🎯 Checking 19 ASR (Attack Surface Reduction) rules...' -ForegroundColor Gray
Write-Verbose 'ASR rules assessed. Starting additional Intune baseline controls assessment.'

# 23. Credential Guard
$cgStatus = $null
try {
    $cgStatus = (Get-WmiObject -Namespace "root\\Microsoft\\Windows\\DeviceGuard" -Class "Win32_DeviceGuard" | Select-Object -ExpandProperty SecurityServicesConfigured) -contains 1
} catch {}
$cgCompliant = $cgStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Credential Guard' `
    -CurrentValue $(if ($cgStatus) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Turn on Credential Guard to protect credentials from theft.' `
    -Remediation 'Enable via GPO or Intune: Turn On Credential Guard' `
    -Documentation 'https://learn.microsoft.com/windows/security/identity-protection/credential-guard/credential-guard' `
    -Compliant $cgCompliant

# 24. Secure Boot
$secureBoot = $null
try {
    $secureBoot = Confirm-SecureBootUEFI
} catch {}
$secureBootCompliant = $secureBoot -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Secure Boot' `
    -CurrentValue $(if ($secureBoot) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Enable Secure Boot in BIOS/UEFI to ensure trusted boot.' `
    -Remediation 'Enable Secure Boot in BIOS/UEFI' `
    -Documentation 'https://learn.microsoft.com/en-us/windows-hardware/design/device-experiences/oem-secure-boot' `
    -Compliant $secureBootCompliant

# 25. WDAC (Application Control)
$wdacStatus = $null
try {
    $wdacStatus = (Get-CimInstance -Namespace "root\\Microsoft\\Windows\\DeviceGuard" -ClassName "Win32_DeviceGuard" | Select-Object -ExpandProperty UserModeCodeIntegrityPolicyEnforced) -eq 1
} catch {}
$wdacCompliant = $wdacStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'WDAC (Application Control)' `
    -CurrentValue $(if ($wdacStatus) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Enable WDAC to restrict execution of unauthorized applications.' `
    -Remediation 'Configure WDAC policy via Intune or GPO' `
    -Documentation 'https://learn.microsoft.com/en-us/windows/security/application-security/application-control/windows-defender-application-control/' `
    -Compliant $wdacCompliant

# 26. Exploit Protection
$exploitProtection = $null
try {
    $exploitProtection = (Get-ProcessMitigation -System | Select-Object -ExpandProperty DEP | Select-Object -ExpandProperty Enable) -eq 1
} catch {}
$exploitCompliant = $exploitProtection -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Exploit Protection (DEP)' `
    -CurrentValue $(if ($exploitProtection) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Keep Exploit Protection (DEP) enabled to mitigate vulnerabilities.' `
    -Remediation 'Configure Exploit Protection via GPO or Intune' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/exploit-protection' `
    -Compliant $exploitCompliant

# 27. SmartScreen
$smartscreenStatus = $null
try {
    $smartscreenStatus = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name SmartScreenEnabled -ErrorAction SilentlyContinue).SmartScreenEnabled -eq 'RequireAdmin'
} catch {}
$smartscreenCompliant = $smartscreenStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'SmartScreen (Legacy)' `
    -CurrentValue $(if ($smartscreenStatus) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Keep SmartScreen enabled to protect against malicious sites and downloads.' `
    -Remediation 'Configure SmartScreen via Intune or GPO' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/microsoft-defender-smartscreen/microsoft-defender-smartscreen-overview' `
    -Compliant $smartscreenCompliant

# 28. Windows Update
$wuStatus = $null
try {
    $wuStatus = (Get-WmiObject -Class "Win32_WindowsUpdateStatus" -ErrorAction SilentlyContinue).UpdateState -eq 1
} catch {}
$wuCompliant = $wuStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Windows Update Automático' `
    -CurrentValue $(if ($wuStatus) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Keep automatic updates enabled to ensure security patches.' `
    -Remediation 'Configure Windows Update via Intune or GPO' `
    -Documentation 'https://learn.microsoft.com/windows/deployment/update/waas-overview' `
    -Compliant $wuCompliant

# 29. UAC
$uacStatus = $null
try {
    $uacStatus = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA -ErrorAction SilentlyContinue).EnableLUA -eq 1
} catch {}
$uacCompliant = $uacStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'UAC (User Account Control)' `
    -CurrentValue $(if ($uacStatus) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Keep UAC enabled to prevent unauthorized elevation.' `
    -Remediation 'Configure UAC via Intune or GPO' `
    -Documentation 'https://learn.microsoft.com/windows/security/identity-protection/user-account-control/how-user-account-control-works' `
    -Compliant $uacCompliant

# 30. Removable Storage Access
$removableStatus = $null
try {
    $removableStatus = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices" -Name Deny_All -ErrorAction SilentlyContinue).Deny_All -eq 1
} catch {}
$removableCompliant = $removableStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Removable Media Access' `
    -CurrentValue $(if ($removableStatus) { 'Restricted' } else { 'Allowed' }) `
    -BestPractice 'Restricted' `
    -Recommendation 'Restrict use of removable media to prevent data leakage.' `
    -Remediation 'Configure restriction via Intune or GPO' `
    -Documentation 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/data-protection/bitlocker/' `
    -Compliant $removableCompliant

# 31. Remote Desktop
$rdpStatus = $null
try {
    $rdpStatus = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name fDenyTSConnections -ErrorAction SilentlyContinue).fDenyTSConnections -eq 1
} catch {}
$rdpCompliant = $rdpStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Remote Desktop Disabled' `
    -CurrentValue $(if ($rdpStatus) { 'Disabled' } else { 'Enabled' }) `
    -BestPractice 'Disabled' `
    -Recommendation 'Disable remote access to reduce attack surface.' `
    -Remediation 'Configure via Intune or GPO: Disable Remote Desktop' `
    -Documentation 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/network-security/windows-firewall/' `
    -Compliant $rdpCompliant

# 32. Account Lockout Policy
$lockoutThreshold = $null
try {
    $lockoutThreshold = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" -Name MaxPasswordAge -ErrorAction SilentlyContinue).MaxPasswordAge
} catch {}
$lockoutCompliant = $null -ne $lockoutThreshold -and $lockoutThreshold -le 30
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Account Lockout Policy' `
    -CurrentValue $(if ($lockoutThreshold) { "$lockoutThreshold days" } else { 'Not configured' }) `
    -BestPractice '≤ 30 days' `
    -Recommendation 'Set account lockout to protect against brute force attacks.' `
    -Remediation 'Configure via Intune or GPO: Account Lockout Policy' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/security-policy-settings/account-lockout-policy' `
    -Compliant $lockoutCompliant

# 33. Audit Policy
$auditStatus = $null
try {
    $auditStatus = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\Security" -Name MaxSize -ErrorAction SilentlyContinue).MaxSize -ge 32768
} catch {}
$auditCompliant = $auditStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Audit Policy (Log ≥ 32MB)' `
    -CurrentValue $(if ($auditStatus) { 'Compliant' } else { 'Not compliant' }) `
    -BestPractice 'Log ≥ 32MB' `
    -Recommendation 'Ensure audit log is sufficient for investigation.' `
    -Remediation 'Configure via Intune or GPO: Audit Policy' `
    -Documentation 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/network-security/windows-firewall/best-practices-configuring' `
    -Compliant $auditCompliant

# 34. Application Guard
$appGuardStatus = $null
try {
    $appGuardStatus = (Get-WindowsOptionalFeature -FeatureName Windows-Defender-ApplicationGuard -Online).State -eq 'Enabled'
} catch {}
$appGuardCompliant = $appGuardStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Application Guard' `
    -CurrentValue $(if ($appGuardStatus) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Turn on Application Guard to isolate untrusted sites.' `
    -Remediation 'Enable via Intune or GPO: Application Guard' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/windows-defender-application-guard/wd-app-guard-overview' `
    -Compliant $appGuardCompliant

# 35. TPM
$tpmStatus = $null
try {
    $tpmStatus = (Get-WmiObject -Namespace "root\\CIMV2\\Security\\MicrosoftTpm" -Class Win32_Tpm -ErrorAction SilentlyContinue).IsActivated_Initial -eq $true
} catch {}
$tpmCompliant = $tpmStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'TPM Enabled' `
    -CurrentValue $(if ($tpmStatus) { 'Enabled' } else { 'Disabled' }) `
    -BestPractice 'Enabled' `
    -Recommendation 'Enable TPM to ensure hardware encryption and integrity.' `
    -Remediation 'Enable TPM via BIOS/UEFI' `
    -Documentation 'https://learn.microsoft.com/windows/security/information-protection/tpm/trusted-platform-module-overview' `
    -Compliant $tpmCompliant

Write-Verbose 'Additional Intune baseline controls assessed. Generating summary and export.'

# 20. Signature last updated
$signatureDate = $mpStatus.AntivirusSignatureLastUpdated
$signatureAge = if ($signatureDate) { (New-TimeSpan -Start $signatureDate -End (Get-Date)).Days } else { [int]::MaxValue }
$signatureCurrent = if ($signatureDate) { $signatureDate.ToString('yyyy-MM-dd HH:mm') } else { 'Unknown' }
$signatureCompliant = $signatureAge -le 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - System Settings' -Setting 'Last signature update date' `
    -CurrentValue $signatureCurrent `
    -BestPractice 'Updated in the last 24h' `
    -Recommendation 'Run Update-MpSignature to keep protection engine up to date.' `
    -Remediation 'Update-MpSignature' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/microsoft-defender-antivirus-updates' `
    -Compliant $signatureCompliant

Write-Host '  🪟 Checking Windows configurations...' -ForegroundColor Gray
Write-Host '  🌐 Checking browser protections...' -ForegroundColor Gray
Write-Host '  ✓  Assessment completed! Generating report...' -ForegroundColor Green
Write-Host

# ---------- Console output ----------

# Custom sorting of categories (4 main groups)
$orderedResults = $assessmentResults | Sort-Object @{Expression={
    switch ($_.Categoria) {
        'MDE Nativo' { 1 }
        'Windows - Configurações de Sistema' { 2 }
        'Compliance' { 3 }
        'Navegadores' { 4 }
        default { 99 }
    }
}}, 'Configuração'

Write-Host
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
$title2 = '🛡️  MICROSOFT DEFENDER FOR ENDPOINT (MDE) ASSESSMENT v2.4 🛡️'
$padLength2 = ($bannerWidth - $title2.Length) / 2
$padLeft2 = ' ' * [math]::Floor($padLength2)
$padRight2 = ' ' * [math]::Ceiling($padLength2)
Write-Host "║$padLeft2$title2$padRight2║" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
$subtitle2 = 'Security Best Practices Compliance Assessment'
$padLengthSub2 = ($bannerWidth - $subtitle2.Length) / 2
$padLeftSub2 = ' ' * [math]::Floor($padLengthSub2)
$padRightSub2 = ' ' * [math]::Ceiling($padLengthSub2)
Write-Host "║$padLeftSub2$subtitle2$padRightSub2║" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
Write-Host

# Overall compliance summary
$totalItems = $orderedResults.Count
$compliantItems = ($orderedResults | Where-Object { $_.Status -eq 'Conforme' }).Count
$nonCompliantItems = $totalItems - $compliantItems
$percentCompliant = if ($totalItems -gt 0) { [math]::Round(($compliantItems / $totalItems) * 100, 2) } else { 0 }

Write-Host
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$scoreTitle = 'OVERALL COMPLIANCE SCORE'
$padLengthScore = ($bannerWidth - $scoreTitle.Length) / 2
$padLeftScore = ' ' * [math]::Floor($padLengthScore)
$padRightScore = ' ' * [math]::Ceiling($padLengthScore)
Write-Host "║$padLeftScore$scoreTitle$padRightScore║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
Write-Host "  📊 Total Items Assessed: " -NoNewline -ForegroundColor White
Write-Host $totalItems -ForegroundColor Yellow
Write-Host "  ✓  Compliant Items: " -NoNewline -ForegroundColor Green
Write-Host $compliantItems -ForegroundColor Green
Write-Host "  ⚠  Attention Items: " -NoNewline -ForegroundColor Yellow
Write-Host $nonCompliantItems -ForegroundColor $(if ($nonCompliantItems -eq 0) { 'Green' } else { 'Red' })
Write-Host
Write-Host "  🎯 OVERALL COMPLIANCE: " -NoNewline -ForegroundColor White
Write-Host "$percentCompliant%" -ForegroundColor $(if ($percentCompliant -ge 80) { 'Green' } elseif ($percentCompliant -ge 60) { 'Yellow' } else { 'Red' }) -NoNewline
Write-Host " ($compliantItems/$totalItems)" -ForegroundColor Gray
Write-Host

# Category compliance summary with enhanced visualization
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$categoryTitle = 'COMPLIANCE SCORE BY CATEGORY'
$padLengthCategory = ($bannerWidth - $categoryTitle.Length) / 2
$padLeftCategory = ' ' * [math]::Floor($padLengthCategory)
$padRightCategory = ' ' * [math]::Ceiling($padLengthCategory)
Write-Host "║$padLeftCategory$categoryTitle$padRightCategory║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host

$categorySummary = $orderedResults | Group-Object Categoria | ForEach-Object {
    $cTotal = $_.Count
    $cConf = ($_.Group | Where-Object { $_.Status -eq 'Conforme' }).Count
    $cNonConf = $cTotal - $cConf
    $cPct = if ($cTotal -gt 0) { [math]::Round(($cConf / $cTotal) * 100, 1) } else { 0 }
    [PSCustomObject]@{ 
        Categoria = $_.Name
        Total = $cTotal
        Conformes = $cConf
        'Não Conformes' = $cNonConf
        'Conformidade' = "$cPct%"
        PctValue = $cPct
    }
} | Sort-Object PctValue -Descending

foreach ($cat in $categorySummary) {
    $color = if ($cat.PctValue -eq 100) { 'Green' } elseif ($cat.PctValue -ge 80) { 'Yellow' } else { 'Red' }
    $icon = if ($cat.PctValue -eq 100) { '✓✓' } elseif ($cat.PctValue -ge 80) { '⚡ ' } else { '⚠⚠' }
    
    Write-Host "  $icon " -NoNewline -ForegroundColor $color
    Write-Host ("{0,-40}" -f $cat.Categoria) -NoNewline -ForegroundColor White
    Write-Host " │ " -NoNewline -ForegroundColor DarkGray
    
    # Visual progress bar
    $barLength = 20
    $filledBars = [math]::Round(($cat.PctValue / 100) * $barLength)
    $emptyBars = $barLength - $filledBars
    Write-Host "[" -NoNewline -ForegroundColor DarkGray
    if ($filledBars -gt 0) {
        Write-Host ([string][char]9608 * $filledBars) -NoNewline -ForegroundColor $color
    }
    if ($emptyBars -gt 0) {
        Write-Host ([string][char]9617 * $emptyBars) -NoNewline -ForegroundColor DarkGray
    }
    Write-Host "] " -NoNewline -ForegroundColor DarkGray
    
    $percentText = "{0,5}%" -f $cat.PctValue
    Write-Host $percentText -NoNewline -ForegroundColor $color
    Write-Host " ($($cat.Conformes)/$($cat.Total))" -ForegroundColor Gray
}
Write-Host

# MDE Native section
Write-Host
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$mdeTitle = '🛡️  MICROSOFT DEFENDER FOR ENDPOINT (MDE) - NATIVE IMPLEMENTATIONS'
$padLengthMde = ($bannerWidth - $mdeTitle.Length) / 2
$padLeftMde = ' ' * [math]::Floor($padLengthMde)
$padRightMde = ' ' * [math]::Ceiling($padLengthMde)
Write-Host "║$padLeftMde$mdeTitle$padRightMde║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
$mdeResults = $orderedResults | Where-Object { $_.Categoria -like 'MDE Nativo*' }
if ($mdeResults) {
    $mdeResults | Select-Object Categoria, 'Configuração', 'Guid', 'Valor Atual', 'Código Atual', 'Código Recomendado', 'Best Practice', Status | Format-Table -AutoSize
}
Write-Host

# Windows section
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$windowsTitle = '🪟  WINDOWS CONFIGURATIONS (non-MDE specific)'
$padLengthWindows = ($bannerWidth - $windowsTitle.Length) / 2
$padLeftWindows = ' ' * [math]::Floor($padLengthWindows)
$padRightWindows = ' ' * [math]::Ceiling($padLengthWindows)
Write-Host "║$padLeftWindows$windowsTitle$padRightWindows║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
$winResults = $orderedResults | Where-Object { $_.Categoria -like 'Windows*' }
if ($winResults) {
    $winResults | Select-Object Categoria, 'Configuração', 'Valor Atual', 'Best Practice', Status | Format-Table -AutoSize
}
Write-Host

# Compliance section
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$complianceTitle = '📋  DEVICE & HARDWARE COMPLIANCE'
$padLengthCompliance = ($bannerWidth - $complianceTitle.Length) / 2
$padLeftCompliance = ' ' * [math]::Floor($padLengthCompliance)
$padRightCompliance = ' ' * [math]::Ceiling($padLengthCompliance)
Write-Host "║$padLeftCompliance$complianceTitle$padRightCompliance║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
$complianceResults = $orderedResults | Where-Object { $_.Categoria -like 'Compliance*' }
if ($complianceResults) {
    $complianceResults | Select-Object Categoria, 'Configuração', 'Valor Atual', 'Best Practice', Status | Format-Table -AutoSize
    Write-Host
}

# Browsers section
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$browserTitle = '🌐  BROWSER PROTECTION'
$padLengthBrowser = ($bannerWidth - $browserTitle.Length) / 2
$padLeftBrowser = ' ' * [math]::Floor($padLengthBrowser)
$padRightBrowser = ' ' * [math]::Ceiling($padLengthBrowser)
Write-Host "║$padLeftBrowser$browserTitle$padRightBrowser║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
$browserResults = $orderedResults | Where-Object { $_.Categoria -like 'Navegadores*' }
if ($browserResults) {
    $browserResults | Select-Object Categoria, 'Configuração', 'Valor Atual', 'Best Practice', Status | Format-Table -AutoSize
    Write-Host
}

# Detailed ASR rules summary
$asrSummary = $orderedResults | Where-Object { $_.Categoria -eq 'MDE Nativo' -and $_.'Configuração' -like '*ASR*' }
if ($asrSummary) {
    Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Cyan
Write-Host ('║{0}║' -f ('🎯 DETAILED ASR RULES SUMMARY (Attack Surface Reduction)'.PadLeft(72).PadRight(83))) -ForegroundColor Cyan
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Cyan
    Write-Host
    $asrSummary | Select-Object Categoria, 'Configuração', 'Guid', 'Código Atual', 'Código Recomendado', 'Valor Atual' | Format-Table -AutoSize
    Write-Host
}

# Items requiring attention
$attentionItems = $orderedResults | Where-Object { $_.Status -ne 'Conforme' }
if ($attentionItems) {
    Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Yellow
Write-Host ('║{0}║' -f ('⚠️  ITEMS REQUIRING IMMEDIATE ATTENTION ⚠️'.PadLeft(62).PadRight(83))) -ForegroundColor Yellow
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Yellow
    Write-Host
    
    $groupedAttention = $attentionItems | Group-Object Categoria | Sort-Object Name
    foreach ($group in $groupedAttention) {
        Write-Host "  📌 $($group.Name)" -ForegroundColor Cyan
        foreach ($item in $group.Group) {
            Write-Host "     ⚠  " -NoNewline -ForegroundColor Red
            Write-Host "$($item.'Configuração')" -NoNewline -ForegroundColor Yellow
            Write-Host " - $($item.'Recomendação')" -ForegroundColor White
            Write-Host "        💡 Fix: " -NoNewline -ForegroundColor DarkGray
            Write-Host "$($item.'Comando de Correção')" -ForegroundColor DarkYellow
        }
        Write-Host
    }
} else {
    Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Green
Write-Host ('║{0}║' -f ('✅ CONGRATULATIONS! ALL COMPLIANT! ✅'.PadLeft(58).PadRight(83))) -ForegroundColor Green
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Green
    Write-Host
    Write-Host '  🎉 All assessed configurations are compliant with best practices!' -ForegroundColor Green
    Write-Host
}

# ---------- Export report ----------
try {
    # Try exporting to XLSX first
    $xlsxPath = [System.IO.Path]::ChangeExtension($OutputPath, 'xlsx')
    $xlsxSuccess = $false
    
    try {
        # Check execution policy
        $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
        if ($currentPolicy -eq 'Restricted') {
            Write-Verbose 'Adjusting execution policy to allow modules...'
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        }

        # Install ImportExcel if needed
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Verbose 'Installing ImportExcel module...'
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
        }
        
        Import-Module ImportExcel -Force
        
        Write-Verbose "Exporting consolidated report to $xlsxPath"
    $checklistSheet = $orderedResults | Select-Object Categoria, 'Configuração', 'Guid', 'Valor Atual', 'Código Atual', 'Código Recomendado', 'Best Practice', 'Recomendação', 'Comando de Correção', 'Documentação', Status
    $recomendacoesSheet = $orderedResults | Where-Object { $_.Status -ne 'Conforme' } | Select-Object Categoria, 'Configuração', 'Valor Atual', 'Best Practice', 'Recomendação', 'Comando de Correção', 'Documentação', Status

        # Export complete checklist
        $checklistSheet | Export-Excel -Path $xlsxPath -WorksheetName 'Checklist' -AutoSize -FreezeTopRow
        
        # Export recommendations to separate tab
        if ($recomendacoesSheet) {
            $recomendacoesSheet | Export-Excel -Path $xlsxPath -WorksheetName 'Recomendacoes' -AutoSize -FreezeTopRow
        }
        
        $xlsxSuccess = $true
        Write-Host
        Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Green
Write-Host ('║{0}║' -f ('✅ XLSX REPORT GENERATED SUCCESSFULLY!'.PadLeft(62).PadRight(83))) -ForegroundColor Green
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Green
        Write-Host
        Write-Host "  📄 File: " -NoNewline -ForegroundColor White
        Write-Host $xlsxPath -ForegroundColor Cyan
        Write-Host "  📊 Created tabs:" -ForegroundColor White
        Write-Host "     • Checklist (All items)" -ForegroundColor Gray
        if ($recomendacoesSheet) {
            Write-Host "     • Recommendations (Non-compliant items)" -ForegroundColor Gray
        }
        Write-Host
        
    } catch {
        Write-Warning "Unable to generate XLSX: $($_.Exception.Message)"
        Write-Warning "Generating fallback in CSV..."
    }
    
    # Fallback to CSV if XLSX fails
    if (-not $xlsxSuccess) {
        Write-Verbose "Exporting consolidated report to $OutputPath"
        $orderedResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        
        # Generate separate CSV for recommendations
        $recomendacoesPath = [System.IO.Path]::ChangeExtension($OutputPath, '_Recomendacoes.csv')
        $recomendacoesSheet = $orderedResults | Where-Object { $_.Status -ne 'Conforme' } | Select-Object Categoria, 'Configuração', 'Valor Atual', 'Best Practice', 'Recomendação', 'Comando de Correção', Status
        if ($recomendacoesSheet) {
            $recomendacoesSheet | Export-Csv -Path $recomendacoesPath -NoTypeInformation -Encoding UTF8
        }
        
        Write-Host
        Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Green
Write-Host ('║{0}║' -f ('✅ CSV REPORT GENERATED SUCCESSFULLY!'.PadLeft(61).PadRight(83))) -ForegroundColor Green
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Green
        Write-Host
        Write-Host "  📄 Main File: " -NoNewline -ForegroundColor White
        Write-Host $OutputPath -ForegroundColor Cyan
        if ($recomendacoesSheet) {
            Write-Host "  📋 Recommendations: " -NoNewline -ForegroundColor White
            Write-Host $recomendacoesPath -ForegroundColor Cyan
        }
        Write-Host
    }
    
} catch {
    Write-Host
    Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Red
Write-Host ('║{0}║' -f ('❌ ERROR EXPORTING REPORT'.PadLeft(58).PadRight(83))) -ForegroundColor Red
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Red
    Write-Host
    Write-Host "  ⚠️  Error: " -NoNewline -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host
    exit 1
}
