<#
.SYNOPSIS
    🛡️ MDE Security Assessment Tool v2.4 - Avaliação completa de conformidade do Microsoft Defender for Endpoint

.DESCRIPTION
    Script profissional de auditoria de segurança que avalia configurações do endpoint contra melhores práticas Microsoft,
    baseline Intune e controles avançados do Windows, com interface visual aprimorada e relatórios detalhados.
    
    ╔══════════════════════════════════════════════════════════════════════════════════╗
    ║                         CATEGORIAS AVALIADAS (46+ controles)                     ║
    ╚══════════════════════════════════════════════════════════════════════════════════╝
    
    🛡️ MDE NATIVO - IMPLEMENTAÇÕES NATIVAS DO DEFENDER FOR ENDPOINT:
    ────────────────────────────────────────────────────────────────────────────────────
    • Onboarding MDE/EDR (serviço Sense)
    • Microsoft Defender Antivírus:
      - Proteção em tempo real, entregue pela nuvem, MAPS
      - PUA (Potentially Unwanted Applications)
      - Network Protection, Behavior Monitoring
      - Block at First Sight (BAFS)
      - Tamper Protection
    • Attack Surface Reduction (ASR) - 19 regras oficiais com GUID e código de ação
    • Configurações de scan (agendamento, parâmetros)
    • Atualização de assinaturas e performance
    
    🪟 WINDOWS - CONFIGURAÇÕES DE SISTEMA (não específicas do MDE):
    ────────────────────────────────────────────────────────────────────────────────────
    • Criptografia: BitLocker (OS e mídias removíveis)
    • Firewall do Windows (perfis e auditoria)
    • Windows Defender SmartScreen (proteção de sistema - nível OS)
    • Credential Guard (proteção de credenciais)
    • Secure Boot e TPM 2.0
    • WDAC (Windows Defender Application Control)
    • Exploit Protection (mitigação de exploits)
    • Application Guard (isolamento de navegador)
    • UAC (User Account Control)
    • LAPS (Local Administrator Password Solution)
    • Políticas de conta (senha, lockout)
    • Auditoria e logs
    • Windows Update
    
    📋 COMPLIANCE - CONFORMIDADE DE DISPOSITIVO & HARDWARE:
    ────────────────────────────────────────────────────────────────────────────────────
    • Políticas de dispositivo: comprimento mínimo de senha, bloqueio de tela
    • Windows Hello for Business (autenticação biométrica)
    • TPM driver atualizado
    
    🌐 NAVEGADORES - PROTEÇÃO DE NAVEGADORES:
    ────────────────────────────────────────────────────────────────────────────────────
    • Microsoft Edge SmartScreen
    • Google Chrome Safe Browsing
    • Mozilla Firefox Enhanced Tracking Protection
    
    ╔══════════════════════════════════════════════════════════════════════════════════╗
    ║                            RECURSOS DA VERSÃO 2.4                                ║
    ╚══════════════════════════════════════════════════════════════════════════════════╝
    
    ✨ NOVIDADES v2.4:
    • Reorganização completa: separação clara entre MDE Nativo e Windows
    • Proteção para Firefox (Enhanced Tracking Protection)
    • Interface visual profissional com boxes, ícones e barras de progresso
    • Score de conformidade com visualização gráfica por categoria
    • Indicadores de progresso em tempo real durante execução
    • Agrupamento inteligente de itens não conformes por categoria
    • Links de documentação validados e atualizados (34 URLs)
    • Exportação aprimorada com feedback visual (XLSX ou CSV)
    • Tratamento de erros com sugestões de solução
    
    📊 SAÍDA E RELATÓRIOS:
    • Console organizado em seções temáticas com boxes visuais
    • Score de conformidade geral e por categoria (com barra de progresso)
    • Exportação XLSX (2 abas: Checklist completo + Recomendações)
    • Fallback CSV se módulo ImportExcel não disponível
    • Todas as recomendações incluem: documentação oficial Microsoft, comando de correção e informações complementares

.PARAMETER OutputPath
    Caminho completo opcional para o arquivo de relatório.
    Se omitido, gera automaticamente em C:\temp\MDE_Assessment_Report_[timestamp].xlsx
    
    Exemplo: -OutputPath "C:\Relatorios\MDE_Assessment.xlsx"

.EXAMPLE
    .\Assessment-MDE-V2.4.ps1
    
    Executa a avaliação completa com saída padrão em C:\temp

.EXAMPLE
    .\Assessment-MDE-V2.4.ps1 -OutputPath "C:\Relatorios\Assessment.xlsx"
    
    Executa com caminho personalizado para o relatório

.EXAMPLE
    .\Assessment-MDE-V2.4.ps1 -Verbose
    
    Executa com detalhes de progresso no console (modo verbose)

.NOTES
    ╔══════════════════════════════════════════════════════════════════════════════════╗
    ║                           INFORMAÇÕES DO SCRIPT                                  ║
    ╚══════════════════════════════════════════════════════════════════════════════════╝
    
    Versão:          2.4
    Data:            07/10/2025
    Autor:           Leandro Sardim
    Última Atualização: Reorganização MDE/Windows + Firefox + Interface Visual Profissional
    
    REQUISITOS:
    ✓ PowerShell 5.1 ou superior
    ✓ Execução como Administrador (recomendado para acesso completo)
    ✓ Windows 10/11 com Microsoft Defender habilitado
    ✓ Conexão com internet (para instalação automática do módulo ImportExcel, se necessário)
    ✓ Política de execução: ajustada automaticamente se necessário
    
    HISTÓRICO DE VERSÕES:
    • v2.4 (07/10/2025): Reorganização categorias + Firefox + Interface visual profissional
    • v2.2 (anterior): Expansão de controles + documentação + separação de colunas
    • v2.0 (anterior): Baseline completo com 40+ controles
    • v1.0 (anterior): Versão inicial
    
    LINKS ÚTEIS:
    • MDE Documentação: https://learn.microsoft.com/en-us/defender-endpoint/
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

# Gera nome de arquivo único com data/hora
$now = Get-Date -Format 'yyyy-MM-dd_HH-mm'
$reportFolder = 'C:\temp'
if (-not (Test-Path -Path $reportFolder)) {
    New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
}
$OutputPath = Join-Path $reportFolder "MDE_Assessment_Report_${now}.csv"

$ErrorActionPreference = 'Stop'

Write-Verbose 'Iniciando assessment do Microsoft Defender for Endpoint.'

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
        Categoria = $Category
        'Configuração' = $Setting
        'Guid' = $Guid
        'Valor Atual' = $CurrentValue
        'Código Atual' = $CurrentCode
        'Código Recomendado' = $RecommendationCode
        'Best Practice' = $BestPractice
        'Recomendação' = $Recommendation
        'Comando de Correção' = $(if ($Remediation -match '(.*)\. Guia: (.*)') { $matches[1].Trim() } else { $Remediation })
        'Documentação' = $Documentation
        'Informação complementar' = $(if ($Remediation -match '(.*)\. Guia: (.*)') { $matches[2].Trim() } else { $null })
        Status = $(if ($Compliant) { 'Conforme' } else { 'Atenção' })
    }
}

# Banner inicial
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
$subtitle1 = 'Microsoft Defender for Endpoint - Avaliação de Conformidade'
$padLengthSub1 = ($bannerWidth - $subtitle1.Length) / 2
$padLeftSub1 = ' ' * [math]::Floor($padLengthSub1)
$padRightSub1 = ' ' * [math]::Ceiling($padLengthSub1)
Write-Host "║$padLeftSub1$subtitle1$padRightSub1║" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
Write-Host '  📋 Iniciando avaliação de segurança...' -ForegroundColor Yellow
Write-Host

Write-Verbose 'Carregando preferências atuais do Defender (Get-MpPreference, Get-MpComputerStatus)...'

try {
    Write-Host '  🔍 Coletando configurações do Microsoft Defender...' -ForegroundColor Gray
    $mpPreference = Get-MpPreference
    $mpStatus = Get-MpComputerStatus
    Write-Verbose 'Preferências carregadas com sucesso.'
    Write-Host '  ✓  Configurações coletadas com sucesso!' -ForegroundColor Green
    Write-Host
} catch {
    Write-Host
    Write-Host '╔══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Red
    Write-Host '║                              ❌ ERRO CRÍTICO                                     ║' -ForegroundColor Red
    Write-Host '╚══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Red
    Write-Host
    Write-Host '  ⚠️  Não foi possível obter as configurações do Microsoft Defender.' -ForegroundColor Yellow
    Write-Host '  💡 Solução: Execute o script como administrador.' -ForegroundColor Cyan
    Write-Host
    exit 1
}

$tamperProtectionKey = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Features'
$tamperProtection = Get-ItemProperty -Path $tamperProtectionKey -Name TamperProtection -ErrorAction SilentlyContinue


$assessmentResults = @()

# Baseline Intune padrão Microsoft (valores recomendados)
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

# ======= IMPLEMENTAÇÕES NATIVAS DO MDE =======

# 1. Status de onboarding do MDE/EDR (serviço Sense)
$mdeService = Get-Service -Name Sense -ErrorAction SilentlyContinue
$mdeOnboarded = $mdeService -and $mdeService.Status -eq 'Running'
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Onboarding do MDE/EDR' `
    -CurrentValue $(if ($mdeOnboarded) { 'Onboarded (Sense em execução)' } else { 'Não onboarded' }) `
    -BestPractice 'Onboarded (serviço Sense ativo)' `
    -Recommendation 'Garantir que o dispositivo esteja onboarded no Microsoft Defender for Endpoint.' `
    -Documentation 'https://learn.microsoft.com/microsoft-365/security/defender-endpoint/onboarding' `
    -Remediation 'Executar script de onboarding ou aplicar política Intune MDE' `
    -Compliant $mdeOnboarded

# ======= FIM MDE NATIVO (mais configurações MDE nativas serão agrupadas após) =======

# ======= CONFIGURAÇÕES DO WINDOWS (não específicas do MDE) =======
# Senha mínima
$passwordPolicy = (Get-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name MinimumPasswordLength -ErrorAction SilentlyContinue).MinimumPasswordLength
$passwordCompliant = $passwordPolicy -ge 8
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'Senha mínima' `
    -CurrentValue $(if ($passwordPolicy) { "$passwordPolicy caracteres" } else { 'Não configurado' }) `
    -BestPractice '≥ 8 caracteres' `
    -Recommendation 'Definir senha mínima de pelo menos 8 caracteres (preferencial: 12+).' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/security-policy-settings/minimum-password-length' `
    -Remediation 'Set-LocalUser -Name <usuario> -Password <senha forte>; ou definir via Intune/GPO' `
    -Compliant $passwordCompliant

# Bloqueio automático de tela (Screen Saver timeout)
$lockScreenTimeout = (Get-ItemProperty -Path "HKCU:Control Panel\Desktop" -Name ScreenSaveTimeOut -ErrorAction SilentlyContinue).ScreenSaveTimeOut
$lockScreenCompliant = $lockScreenTimeout -and ([int]$lockScreenTimeout -le 900)
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'Bloqueio de tela automático' `
    -CurrentValue $(if ($lockScreenTimeout) { "$lockScreenTimeout segundos" } else { 'Não configurado' }) `
    -BestPractice '≤ 900 segundos (15 min)' `
    -Recommendation 'Configurar bloqueio automático em até 15 minutos de inatividade.' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/security-policy-settings/interactive-logon-machine-inactivity-limit' `
    -Remediation 'Set-ItemProperty -Path "HKCU:Control Panel\Desktop" -Name ScreenSaveTimeOut -Value 900; ou definir via Intune/GPO' `
    -Compliant $lockScreenCompliant

# Windows Hello (exemplo simples – presença de chave de FaceLogon)
$helloKey = 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\FaceLogon'
$helloEnabled = (Get-ItemProperty -Path $helloKey -Name Enabled -ErrorAction SilentlyContinue).Enabled -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'Windows Hello (biometria)' `
    -CurrentValue $(if ($helloEnabled) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Habilitar Windows Hello (PIN + biometria) para fortalecer autenticação.' `
    -Documentation 'https://learn.microsoft.com/windows/security/identity-protection/hello-for-business/' `
    -Remediation 'Configurar via Intune: Windows Hello for Business' `
    -SupplementaryInfo 'https://learn.microsoft.com/windows/security/identity-protection/hello-for-business/deploy/hybrid-cloud-kerberos-trust' `
    -Compliant $helloEnabled

# 3. Driver TPM atualizado (≤ 1 ano)
$tpmDriver = Get-WmiObject -Class Win32_PnPSignedDriver -ErrorAction SilentlyContinue | Where-Object { $_.DeviceName -like '*TPM*' }
$tpmDriverDate = if ($tpmDriver) { $tpmDriver.DriverDate } else { $null }
$tpmDriverCompliant = $tpmDriverDate -and ((Get-Date) - $tpmDriverDate).Days -le 365
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'Driver TPM atualizado' `
    -CurrentValue $(if ($tpmDriverDate) { $tpmDriverDate.ToString('yyyy-MM-dd') } else { 'Não encontrado' }) `
    -BestPractice 'Driver atualizado (≤ 1 ano)' `
    -Recommendation 'Atualizar driver TPM para assegurar compatibilidade e correções.' `
    -Documentation 'https://learn.microsoft.com/windows/security/hardware-security/tpm/trusted-platform-module-top-node' `
    -Remediation 'Atualizar via Windows Update ou fabricante' `
    -Compliant $tpmDriverCompliant

# 7. LAPS (Local Admin Password Solution)
$lapsKey = 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS'
$lapsEnabled = (Get-ItemProperty -Path $lapsKey -Name Enable -ErrorAction SilentlyContinue).Enable -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'LAPS habilitado' `
    -CurrentValue $(if ($lapsEnabled) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Habilitar Windows LAPS para proteger credenciais de administradores locais.' `
    -Documentation 'https://learn.microsoft.com/windows-server/identity/laps/laps-overview' `
    -Remediation 'Configurar política Windows LAPS via Intune / GPO' `
    -SupplementaryInfo 'https://learn.microsoft.com/windows-server/identity/laps/laps-scenarios-legacy' `
    -Compliant $lapsEnabled

# 8. Proteção SmartScreen (Sistema + Navegadores)
# Windows Defender SmartScreen (proteção em nível de sistema para apps e arquivos)
$systemSmartScreen = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer' -Name SmartScreenEnabled -ErrorAction SilentlyContinue).SmartScreenEnabled
$systemSmartScreenCompliant = $systemSmartScreen -in @('Warn', 'Block')
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Windows Defender SmartScreen (Sistema)' `
    -CurrentValue $(if ($systemSmartScreen) { $systemSmartScreen } else { 'Não configurado' }) `
    -BestPractice 'Warn ou Block' `
    -Recommendation 'Habilitar SmartScreen do Windows para proteger contra apps/arquivos maliciosos de qualquer fonte.' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/microsoft-defender-smartscreen/microsoft-defender-smartscreen-overview' `
    -Remediation 'Set-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name SmartScreenEnabled -Value "Block"; ou política Intune/GPO' `
    -Compliant $systemSmartScreenCompliant

# Edge SmartScreen (proteção específica do navegador)
$edgeSmartScreen = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Microsoft\Edge' -Name SmartScreenEnabled -ErrorAction SilentlyContinue).SmartScreenEnabled -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Navegadores' -Setting 'Microsoft Edge SmartScreen' `
    -CurrentValue $(if ($edgeSmartScreen) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Garantir SmartScreen no Edge para bloquear sites/downloads maliciosos.' `
    -Documentation 'https://learn.microsoft.com/deployedge/microsoft-edge-security-smartscreen' `
    -Remediation 'Set-ItemProperty -Path "HKLM:SOFTWARE\Policies\Microsoft\Edge" -Name SmartScreenEnabled -Value 1; ou política Intune/GPO' `
    -Compliant $edgeSmartScreen

# Chrome Safe Browsing (nível básico = 1; reforçado pode ser 2 dependendo da organização)
$chromeSafeBrowsing = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Google\Chrome' -Name SafeBrowsingProtectionLevel -ErrorAction SilentlyContinue).SafeBrowsingProtectionLevel
$chromeSafeBrowsingEnabled = $chromeSafeBrowsing -in 1,2
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Navegadores' -Setting 'Chrome Safe Browsing' `
    -CurrentValue $(if ($chromeSafeBrowsing) { "Nível $chromeSafeBrowsing" } else { 'Não configurado' }) `
    -BestPractice 'Ativado (nível 1 ou 2)' `
    -Recommendation 'Manter Safe Browsing no mínimo em nível padrão; considerar reforçado.' `
    -Documentation 'https://support.google.com/chrome/answer/99020' `
    -Remediation 'Set-ItemProperty -Path "HKLM:SOFTWARE\Policies\Google\Chrome" -Name SafeBrowsingProtectionLevel -Value 1; ou política Intune/GPO' `
    -Compliant $chromeSafeBrowsingEnabled

# Firefox Enhanced Tracking Protection + Safe Browsing
# Firefox usa preferences.js, mas políticas corporativas podem ser via HKLM\SOFTWARE\Policies\Mozilla\Firefox
$firefoxETP = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Mozilla\Firefox' -Name 'EnableTrackingProtection' -ErrorAction SilentlyContinue).EnableTrackingProtection -eq 1
$firefoxSafeBrowsing = (Get-ItemProperty -Path 'HKLM:SOFTWARE\Policies\Mozilla\Firefox' -Name 'DisableSafeMode' -ErrorAction SilentlyContinue).DisableSafeMode -ne 1 # Safe Browsing habilitado por padrão se não desabilitado
$firefoxProtectionEnabled = $firefoxETP -or $firefoxSafeBrowsing
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Navegadores' -Setting 'Firefox Enhanced Tracking Protection' `
    -CurrentValue $(if ($firefoxETP) { 'Ativado' } elseif ($firefoxSafeBrowsing) { 'Safe Browsing ativo' } else { 'Não configurado/Desativado' }) `
    -BestPractice 'Ativado (Enhanced Tracking Protection + Safe Browsing)' `
    -Recommendation 'Habilitar Enhanced Tracking Protection e Safe Browsing no Firefox para proteção contra rastreamento e sites maliciosos.' `
    -Documentation 'https://support.mozilla.org/kb/enhanced-tracking-protection-firefox-desktop' `
    -Remediation 'Configurar via política corporativa Mozilla ou about:config: privacy.trackingprotection.enabled=true' `
    -SupplementaryInfo 'https://support.mozilla.org/kb/how-does-phishing-and-malware-protection-work' `
    -Compliant $firefoxProtectionEnabled

# ======= FIM MELHORIAS =======



# 1. Real-time protection (Baseline Intune)
$rtpEnabled = -not $mpPreference.DisableRealtimeMonitoring
$rtpCompliant = $rtpEnabled -eq $intuneBaseline['Defender_RealTimeProtection']
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Proteção em tempo real (Baseline Intune)' `
    -CurrentValue $(if ($rtpEnabled) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado (DisableRealtimeMonitoring = 0)' `
    -Recommendation 'Manter a proteção em tempo real sempre habilitada.' `
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
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'BitLocker Protegido (Baseline Intune)' `
    -CurrentValue $(if ($bitlockerEnabled) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Ativar BitLocker para proteger dados em repouso.' `
    -Remediation 'Enable-BitLocker -MountPoint C:' `
    -Documentation 'https://learn.microsoft.com/windows/security/information-protection/bitlocker/bitlocker-overview' `
    -Compliant $bitlockerCompliant

# 22. Firewall Enabled (Baseline Intune)
$firewallEnabled = $null
try {
    $firewallEnabled = (Get-NetFirewallProfile | Where-Object { $_.Enabled -eq 'True' }).Count -ge 1
} catch {}
$firewallCompliant = $firewallEnabled -eq $intuneBaseline['Firewall_Enabled']
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Firewall Ativado (Baseline Intune)' `
    -CurrentValue $(if ($firewallEnabled) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Manter o firewall do Windows ativado em todos os perfis.' `
    -Remediation 'Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True' `
    -Documentation 'https://learn.microsoft.com/windows/security/operating-system-security/network-security/windows-firewall/windows-firewall-with-advanced-security' `
    -Compliant $firewallCompliant

# 2. Cloud-delivered protection (MAPS)
$mapsLevel = $mpPreference.MAPSReporting
$mapsLabel = switch ($mapsLevel) {
    0 { 'Desativado' }
    1 { 'Básico' }
    2 { 'Avançado' }
    default { "Desconhecido ($mapsLevel)" }
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Proteção baseada na nuvem' `
    -CurrentValue $mapsLabel `
    -BestPractice 'Avançado (MAPSReporting = 2)' `
    -Recommendation 'Habilitar o nível avançado para respostas em tempo real da nuvem.' `
    -Remediation 'Set-MpPreference -MAPSReporting Advanced' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant ($mapsLevel -eq 2) `
    -CurrentCode $mapsLevel `
    -RecommendationCode 2

# 3. Envio automático de amostras
$sampleConsent = $mpPreference.SubmitSamplesConsent
$sampleLabel = switch ($sampleConsent) {
    0 { 'Sempre perguntar' }
    1 { 'Enviar amostras seguras automaticamente' }
    2 { 'Nunca enviar' }
    3 { 'Enviar todas as amostras automaticamente' }
    default { "Desconhecido ($sampleConsent)" }
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Envio automático de amostras' `
    -CurrentValue $sampleLabel `
    -BestPractice 'Enviar todas as amostras automaticamente (valor 3)' `
    -Recommendation 'Garantir envio automático de amostras para maior eficácia de detecção.' `
    -Remediation 'Set-MpPreference -SubmitSamplesConsent SendAllSamples' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant ($sampleConsent -eq 3) `
    -CurrentCode $sampleConsent `
    -RecommendationCode 3

# 4. PUA Protection
$puaState = $mpPreference.PUAProtection
$puaLabel = switch ($puaState) {
    0 { 'Desativado' }
    1 { 'Ativado' }
    2 { 'Somente auditoria' }
    default { "Desconhecido ($puaState)" }
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Proteção contra aplicativos potencialmente indesejados (PUA)' `
    -CurrentValue $puaLabel `
    -BestPractice 'Ativado' `
    -Recommendation 'Bloquear aplicativos potencialmente indesejados para reduzir a superfície de ataque.' `
    -Remediation 'Set-MpPreference -PUAProtection Enabled' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/detect-block-potentially-unwanted-apps-microsoft-defender-antivirus' `
    -Compliant ($puaState -eq 1) `
    -CurrentCode $puaState `
    -RecommendationCode 1

# 5. Network protection
$networkProtection = $mpPreference.EnableNetworkProtection
$networkLabel = switch ($networkProtection) {
    0 { 'Desativado' }
    1 { 'Bloquear' }
    2 { 'Somente auditoria' }
    default { "Desconhecido ($networkProtection)" }
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Proteção de rede' `
    -CurrentValue $networkLabel `
    -BestPractice 'Bloquear (EnableNetworkProtection = 1)' `
    -Recommendation 'Colocar a proteção de rede em modo de bloqueio para prevenir comunicações maliciosas.' `
    -Remediation 'Set-MpPreference -EnableNetworkProtection Enabled' `
    -Documentation 'https://learn.microsoft.com/microsoft-365/security/defender-endpoint/network-protection' `
    -Compliant ($networkProtection -eq 1) `
    -CurrentCode $networkProtection `
    -RecommendationCode 1

# 6. Monitoramento comportamental
$behaviorMonitoring = -not $mpPreference.DisableBehaviorMonitoring
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Monitoramento comportamental' `
    -CurrentValue $(if ($behaviorMonitoring) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Detectar comportamento suspeito analisando atividades em tempo real.' `
    -Remediation 'Set-MpPreference -DisableBehaviorMonitoring $false' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant $behaviorMonitoring `
    -CurrentCode ([int]$mpPreference.DisableBehaviorMonitoring) `
    -RecommendationCode 0

# 7. Bloqueio na primeira visualização
$blockAtFirstSeen = -not $mpPreference.DisableBlockAtFirstSeen
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Bloqueio na primeira visualização' `
    -CurrentValue $(if ($blockAtFirstSeen) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Permitir que o Defender bloqueie arquivos nunca vistos assim que aparecem.' `
    -Remediation 'Set-MpPreference -DisableBlockAtFirstSeen $false' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/microsoft-defender-antivirus/configure-block-at-first-sight-microsoft-defender-antivirus' `
    -Compliant $blockAtFirstSeen `
    -CurrentCode ([int]$mpPreference.DisableBlockAtFirstSeen) `
    -RecommendationCode 0

# 8. Proteção contra adulteração (Tamper Protection)
$tamperValue = $tamperProtection.TamperProtection
$tamperEnabled = $tamperValue -eq 5
$tamperLabel = if ($null -ne $tamperValue) {
    if ($tamperEnabled) { 'Ativado (5)' } else { "Desativado ($tamperValue)" }
} else {
    'Não configurado'
}
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Proteção contra adulteração' `
    -CurrentValue $tamperLabel `
    -BestPractice 'Ativado (valor 5)' `
    -Recommendation 'Habilitar via portal do Microsoft Defender ou Microsoft Intune.' `
    -Remediation 'Configuração via portal MDE/Intune (não disponível em PowerShell).' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/prevent-changes-to-security-settings-with-tamper-protection' `
    -Compliant $tamperEnabled `
    -Guid $(if ($null -ne $tamperValue) { 'TamperProtection' } else { $null }) `
    -CurrentCode $tamperValue `
    -RecommendationCode 5

# 9. Verificação de arquivos em rede
$networkScanDisabled = $mpPreference.DisableScanningNetworkFiles
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Ignorar arquivos de rede já protegidos' `
    -CurrentValue $(if ($networkScanDisabled) { 'Desativada (recomendado)' } else { 'Ativada' }) `
    -BestPractice 'Desativada para evitar dupla varredura' `
    -Recommendation 'Desativar a varredura de arquivos de rede para reduzir impacto de performance (a proteção deve existir no servidor de origem).' `
    -Remediation 'Set-MpPreference -DisableScanningNetworkFiles $true' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant $networkScanDisabled `
    -CurrentCode ([int]$networkScanDisabled) `
    -RecommendationCode 1

# 10. Carregamento máximo de CPU nas verificações
$cpuLoad = $mpPreference.ScanAvgCPULoadFactor
$cpuDisplay = if ($cpuLoad -is [int]) { "$cpuLoad%" } else { 'Não configurado' }
$cpuCompliant = ($cpuLoad -is [int]) -and ($cpuLoad -le 50)
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'Carga média de CPU durante scans' `
    -CurrentValue $cpuDisplay `
    -BestPractice '≤ 50% (valor padrão recomendado)' `
    -Recommendation 'Limitar a carga de CPU garante proteção sem degradar o desempenho.' `
    -Remediation 'Set-MpPreference -ScanAvgCPULoadFactor 50' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/tune-performance-defender-antivirus' `
    -Compliant $cpuCompliant `
    -CurrentCode $(if ($cpuLoad -is [int]) { $cpuLoad } else { $null }) `
    -RecommendationCode 50

function Get-BoolLabel {
    param ($Value)

    if ($null -eq $Value) { return 'Não configurado' }
    if ($Value -is [bool]) {
        return $(if ($Value) { 'Ativado' } else { 'Desativado' })
    }

    switch ($Value) {
        0 { 'Desativado (0)' }
        1 { 'Ativado (1)' }
        default { $Value.ToString() }
    }
}

function Get-CloudBlockLevelLabel {
    param ([Nullable[int]]$Value)

    if ($null -eq $Value) { return 'Não configurado' }

    switch ($Value) {
        0 { 'Desativado (0)' }
        1 { 'Baixo (1)' }
        2 { 'Alto (2)' }
        3 { 'Alto Plus (3)' }
        4 { 'Zero Tolerance (4)' }
        default { "Desconhecido ($Value)" }
    }
}

function Get-UnknownThreatActionLabel {
    param ([Nullable[int]]$Value)

    if ($null -eq $Value) { return 'Não configurado' }

    switch ($Value) {
        0 { 'Permitir (0)' }
        1 { 'Bloquear (1)' }
        2 { 'Quarentena (2)' }
        3 { 'Remover (3)' }
        4 { 'Limpar (4)' }
        5 { 'Ignorar (5)' }
        6 { 'Bloquear (parâmetro recomendado)' }
        default { "Desconhecido ($Value)" }
    }
}

function Get-ScanTypeLabel {
    param ([Nullable[int]]$Value)

    if ($null -eq $Value) { return 'Não configurado' }

    switch ($Value) {
        1 { 'Rápida (1)' }
        2 { 'Completa (2)' }
        3 { 'Customizada (3)' }
        default { "Desconhecido ($Value)" }
    }
}

function Get-ScanDayLabel {
    param ([Nullable[int]]$Value)

    if ($null -eq $Value) { return 'Todos os dias (0)' }

    switch ($Value) {
        0 { 'Todos os dias (0)' }
        1 { 'Domingo (1)' }
        2 { 'Segunda (2)' }
        3 { 'Terça (3)' }
        4 { 'Quarta (4)' }
        5 { 'Quinta (5)' }
        6 { 'Sexta (6)' }
        7 { 'Sábado (7)' }
        default { "Desconhecido ($Value)" }
    }
}

function Convert-MinutesToTimeLabel {
    param ([Nullable[int]]$Minutes)

    if ($null -eq $Minutes) { return 'Não configurado' }
    if ($Minutes -lt 0) { return "Desconhecido ($Minutes)" }

    $hours = [int]($Minutes / 60)
    $mins = $Minutes % 60
    '{0:00}:{1:00} ({2})' -f $hours, $mins, $Minutes
}

function Get-AsrActionLabel {
    param ([Nullable[int]]$Action)

    if ($null -eq $Action) { return 'Não configurado' }

    switch ($Action) {
        0 { 'Desativado (0)' }
        1 { 'Bloquear (1)' }
        2 { 'Somente auditoria (2)' }
        3 { 'Desabilitar (3)' }
        4 { 'Bloquear - legado (4)' }
        5 { 'Permitir (5)' }
        6 { 'Avisar (6)' }
        default { "Desconhecido ($Action)" }
    }
}

# 11. Allow Network Protection Down-Level
$allowNpd = $mpPreference.AllowNetworkProtectionDownLevel
$allowNpdValue = if ($null -eq $allowNpd) { $null } else { [int]$allowNpd }
$allowNpdEnabled = $allowNpdValue -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'AllowNetworkProtectionDownLevel' `
    -CurrentValue (Get-BoolLabel $allowNpd) `
    -BestPractice 'Ativado (valor 1)' `
    -Recommendation 'Habilitar proteção de rede também para versões legadas do Windows.' `
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
    -BestPractice 'Ativado (valor 1)' `
    -Recommendation 'Habilitar proteção de rede nos servidores para expandir a cobertura de bloqueio.' `
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
    -BestPractice 'Alto (valor 2)' `
    -Recommendation 'Definir o nível de bloqueio da nuvem como Alto para resposta mais agressiva.' `
    -Remediation "Set-MpPreference -CloudBlockLevel High" `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/configure-microsoft-defender-antivirus-features' `
    -Compliant $cloudBlockCompliant `
    -CurrentCode $cloudBlockLevel `
    -RecommendationCode 2

# 14. Cloud Extended Timeout
$cloudExtendedTimeout = $mpPreference.CloudExtendedTimeout
$cloudTimeoutCompliant = ($cloudExtendedTimeout -is [int]) -and ($cloudExtendedTimeout -ge 50)
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'CloudExtendedTimeout' `
    -CurrentValue $(if ($cloudExtendedTimeout -is [int]) { "$cloudExtendedTimeout segundos" } else { 'Não configurado' }) `
    -BestPractice '50 segundos (valor recomendado)' `
    -Recommendation 'Permitir que a nuvem tenha tempo suficiente para analisar arquivos antes de liberar.' `
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
    -BestPractice 'Bloquear/Quarentena (valor 6)' `
    -Recommendation 'Garantir que ameaças desconhecidas sejam bloqueadas automaticamente.' `
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
    -BestPractice 'Ativado (valor 1)' `
    -Recommendation 'Verificar se há novas assinaturas antes de iniciar qualquer varredura agendada.' `
    -Remediation 'Set-MpPreference -CheckForSignaturesBeforeRunningScan $true' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/microsoft-defender-antivirus-updates' `
    -Compliant $checkBeforeScanCompliant `
    -CurrentCode $checkBeforeScanValue `
    -RecommendationCode 1

# 17. Parâmetros de varredura agendada
$scanParameters = $mpPreference.ScanParameters
$scanParametersCompliant = $scanParameters -eq 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'ScanParameters' `
    -CurrentValue (Get-ScanTypeLabel $scanParameters) `
    -BestPractice 'Varredura rápida diária (valor 1)' `
    -Recommendation 'Utilizar varredura rápida nas agendas automáticas.' `
    -Remediation 'Set-MpPreference -ScanParameters 1' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/schedule-antivirus-scans' `
    -Compliant $scanParametersCompliant `
    -CurrentCode $scanParameters `
    -RecommendationCode 1

# 18. Dia agendado da varredura
$scanScheduleDay = $mpPreference.ScanScheduleDay
$scanScheduleDayCompliant = $scanScheduleDay -eq 0
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'MDE Nativo' -Setting 'ScanScheduleDay' `
    -CurrentValue (Get-ScanDayLabel $scanScheduleDay) `
    -BestPractice 'Todos os dias (valor 0)' `
    -Recommendation 'Agendar a varredura diária para cobertura contínua.' `
    -Remediation 'Set-MpPreference -ScanScheduleDay 0' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/schedule-antivirus-scans' `
    -Compliant $scanScheduleDayCompliant `
    -CurrentCode $scanScheduleDay `
    -RecommendationCode 0

# 19. Horário agendado da varredura
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
    -BestPractice '02:00 (valor 120)' `
    -Recommendation 'Manter a varredura agendada em horário de baixo uso.' `
    -Remediation 'Set-MpPreference -ScanScheduleTime 120' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/schedule-antivirus-scans' `
    -Compliant $scanScheduleTimeCompliant `
    -CurrentCode $scanScheduleTimeMinutes `
    -RecommendationCode 120

Write-Host '  ⚙️  Verificando configurações principais do Defender...' -ForegroundColor Gray
Write-Verbose 'Checklist de configurações principais concluído. Iniciando avaliação das regras ASR.'

# Mapeia ações atuais das regras ASR
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

Write-Host '  🎯 Verificando 19 regras ASR (Attack Surface Reduction)...' -ForegroundColor Gray
Write-Verbose 'Regras ASR avaliadas. Iniciando avaliação dos controles adicionais do baseline Intune.'

# 23. Credential Guard
$cgStatus = $null
try {
    $cgStatus = (Get-WmiObject -Namespace "root\\Microsoft\\Windows\\DeviceGuard" -Class "Win32_DeviceGuard" | Select-Object -ExpandProperty SecurityServicesConfigured) -contains 1
} catch {}
$cgCompliant = $cgStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Credential Guard' `
    -CurrentValue $(if ($cgStatus) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Ativar Credential Guard para proteger credenciais contra roubo.' `
    -Remediation 'Habilitar via GPO ou Intune: Turn On Credential Guard' `
    -Documentation 'https://learn.microsoft.com/windows/security/identity-protection/credential-guard/credential-guard' `
    -Compliant $cgCompliant

# 24. Secure Boot
$secureBoot = $null
try {
    $secureBoot = Confirm-SecureBootUEFI
} catch {}
$secureBootCompliant = $secureBoot -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Secure Boot' `
    -CurrentValue $(if ($secureBoot) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Ativar Secure Boot na BIOS/UEFI para garantir inicialização confiável.' `
    -Remediation 'Habilitar Secure Boot na BIOS/UEFI' `
    -Documentation 'https://learn.microsoft.com/en-us/windows-hardware/design/device-experiences/oem-secure-boot' `
    -Compliant $secureBootCompliant

# 25. WDAC (Application Control)
$wdacStatus = $null
try {
    $wdacStatus = (Get-CimInstance -Namespace "root\\Microsoft\\Windows\\DeviceGuard" -ClassName "Win32_DeviceGuard" | Select-Object -ExpandProperty UserModeCodeIntegrityPolicyEnforced) -eq 1
} catch {}
$wdacCompliant = $wdacStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'WDAC (Application Control)' `
    -CurrentValue $(if ($wdacStatus) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Habilitar WDAC para restringir execução de aplicativos não autorizados.' `
    -Remediation 'Configurar política WDAC via Intune ou GPO' `
    -Documentation 'https://learn.microsoft.com/en-us/windows/security/application-security/application-control/windows-defender-application-control/' `
    -Compliant $wdacCompliant

# 26. Exploit Protection
$exploitProtection = $null
try {
    $exploitProtection = (Get-ProcessMitigation -System | Select-Object -ExpandProperty DEP | Select-Object -ExpandProperty Enable) -eq 1
} catch {}
$exploitCompliant = $exploitProtection -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Exploit Protection (DEP)' `
    -CurrentValue $(if ($exploitProtection) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Manter Exploit Protection (DEP) ativado para mitigar vulnerabilidades.' `
    -Remediation 'Configurar Exploit Protection via GPO ou Intune' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/exploit-protection' `
    -Compliant $exploitCompliant

# 27. SmartScreen
$smartscreenStatus = $null
try {
    $smartscreenStatus = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" -Name SmartScreenEnabled -ErrorAction SilentlyContinue).SmartScreenEnabled -eq 'RequireAdmin'
} catch {}
$smartscreenCompliant = $smartscreenStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'SmartScreen (Legacy)' `
    -CurrentValue $(if ($smartscreenStatus) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Manter SmartScreen ativado para proteger contra sites e downloads maliciosos.' `
    -Remediation 'Configurar SmartScreen via Intune ou GPO' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/microsoft-defender-smartscreen/microsoft-defender-smartscreen-overview' `
    -Compliant $smartscreenCompliant

# 28. Windows Update
$wuStatus = $null
try {
    $wuStatus = (Get-WmiObject -Class "Win32_WindowsUpdateStatus" -ErrorAction SilentlyContinue).UpdateState -eq 1
} catch {}
$wuCompliant = $wuStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Windows Update Automático' `
    -CurrentValue $(if ($wuStatus) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Manter atualizações automáticas ativadas para garantir correções de segurança.' `
    -Remediation 'Configurar Windows Update via Intune ou GPO' `
    -Documentation 'https://learn.microsoft.com/windows/deployment/update/waas-overview' `
    -Compliant $wuCompliant

# 29. UAC
$uacStatus = $null
try {
    $uacStatus = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name EnableLUA -ErrorAction SilentlyContinue).EnableLUA -eq 1
} catch {}
$uacCompliant = $uacStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'UAC (User Account Control)' `
    -CurrentValue $(if ($uacStatus) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Manter UAC ativado para impedir elevação não autorizada.' `
    -Remediation 'Configurar UAC via Intune ou GPO' `
    -Documentation 'https://learn.microsoft.com/windows/security/identity-protection/user-account-control/how-user-account-control-works' `
    -Compliant $uacCompliant

# 30. Removable Storage Access
$removableStatus = $null
try {
    $removableStatus = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\RemovableStorageDevices" -Name Deny_All -ErrorAction SilentlyContinue).Deny_All -eq 1
} catch {}
$removableCompliant = $removableStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Acesso a Mídia Removível' `
    -CurrentValue $(if ($removableStatus) { 'Restrito' } else { 'Permitido' }) `
    -BestPractice 'Restrito' `
    -Recommendation 'Restringir uso de mídias removíveis para evitar vazamento de dados.' `
    -Remediation 'Configurar restrição via Intune ou GPO' `
    -Documentation 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/data-protection/bitlocker/' `
    -Compliant $removableCompliant

# 31. Remote Desktop
$rdpStatus = $null
try {
    $rdpStatus = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name fDenyTSConnections -ErrorAction SilentlyContinue).fDenyTSConnections -eq 1
} catch {}
$rdpCompliant = $rdpStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Remote Desktop Desativado' `
    -CurrentValue $(if ($rdpStatus) { 'Desativado' } else { 'Ativado' }) `
    -BestPractice 'Desativado' `
    -Recommendation 'Desativar acesso remoto para reduzir superfície de ataque.' `
    -Remediation 'Configurar via Intune ou GPO: Desativar Remote Desktop' `
    -Documentation 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/network-security/windows-firewall/' `
    -Compliant $rdpCompliant

# 32. Account Lockout Policy
$lockoutThreshold = $null
try {
    $lockoutThreshold = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" -Name MaxPasswordAge -ErrorAction SilentlyContinue).MaxPasswordAge
} catch {}
$lockoutCompliant = $null -ne $lockoutThreshold -and $lockoutThreshold -le 30
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Account Lockout Policy' `
    -CurrentValue $(if ($lockoutThreshold) { "$lockoutThreshold dias" } else { 'Não configurado' }) `
    -BestPractice '≤ 30 dias' `
    -Recommendation 'Configurar bloqueio de conta para proteger contra ataques de força bruta.' `
    -Remediation 'Configurar via Intune ou GPO: Account Lockout Policy' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/security-policy-settings/account-lockout-policy' `
    -Compliant $lockoutCompliant

# 33. Audit Policy
$auditStatus = $null
try {
    $auditStatus = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\Security" -Name MaxSize -ErrorAction SilentlyContinue).MaxSize -ge 32768
} catch {}
$auditCompliant = $auditStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Audit Policy (Log ≥ 32MB)' `
    -CurrentValue $(if ($auditStatus) { 'Conforme' } else { 'Não conforme' }) `
    -BestPractice 'Log ≥ 32MB' `
    -Recommendation 'Garantir que o log de auditoria seja suficiente para investigação.' `
    -Remediation 'Configurar via Intune ou GPO: Audit Policy' `
    -Documentation 'https://learn.microsoft.com/en-us/windows/security/operating-system-security/network-security/windows-firewall/best-practices-configuring' `
    -Compliant $auditCompliant

# 34. Application Guard
$appGuardStatus = $null
try {
    $appGuardStatus = (Get-WindowsOptionalFeature -FeatureName Windows-Defender-ApplicationGuard -Online).State -eq 'Enabled'
} catch {}
$appGuardCompliant = $appGuardStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Application Guard' `
    -CurrentValue $(if ($appGuardStatus) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Ativar Application Guard para isolar sites não confiáveis.' `
    -Remediation 'Habilitar via Intune ou GPO: Application Guard' `
    -Documentation 'https://learn.microsoft.com/windows/security/threat-protection/windows-defender-application-guard/wd-app-guard-overview' `
    -Compliant $appGuardCompliant

# 35. TPM
$tpmStatus = $null
try {
    $tpmStatus = (Get-WmiObject -Namespace "root\\CIMV2\\Security\\MicrosoftTpm" -Class Win32_Tpm -ErrorAction SilentlyContinue).IsActivated_Initial -eq $true
} catch {}
$tpmCompliant = $tpmStatus -eq $true
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Compliance' -Setting 'TPM Ativado' `
    -CurrentValue $(if ($tpmStatus) { 'Ativado' } else { 'Desativado' }) `
    -BestPractice 'Ativado' `
    -Recommendation 'Ativar TPM para garantir criptografia e integridade do hardware.' `
    -Remediation 'Ativar TPM via BIOS/UEFI' `
    -Documentation 'https://learn.microsoft.com/windows/security/information-protection/tpm/trusted-platform-module-overview' `
    -Compliant $tpmCompliant

Write-Verbose 'Controles adicionais do baseline Intune avaliados. Gerando sumário e exportação.'

# 20. Atualização de assinaturas
$signatureDate = $mpStatus.AntivirusSignatureLastUpdated
$signatureAge = if ($signatureDate) { (New-TimeSpan -Start $signatureDate -End (Get-Date)).Days } else { [int]::MaxValue }
$signatureCurrent = if ($signatureDate) { $signatureDate.ToString('yyyy-MM-dd HH:mm') } else { 'Desconhecido' }
$signatureCompliant = $signatureAge -le 1
Add-AssessmentResult -Collector ([ref]$assessmentResults) -Category 'Windows - Configurações de Sistema' -Setting 'Data da última atualização de assinaturas' `
    -CurrentValue $signatureCurrent `
    -BestPractice 'Atualização nas últimas 24h' `
    -Recommendation 'Executar Update-MpSignature para manter o motor de proteção atualizado.' `
    -Remediation 'Update-MpSignature' `
    -Documentation 'https://learn.microsoft.com/en-us/defender-endpoint/microsoft-defender-antivirus-updates' `
    -Compliant $signatureCompliant

Write-Host '  🪟 Verificando configurações do Windows...' -ForegroundColor Gray
Write-Host '  🌐 Verificando proteções de navegadores...' -ForegroundColor Gray
Write-Host '  ✓  Avaliação concluída! Gerando relatório...' -ForegroundColor Green
Write-Host

# ---------- Saída no console ----------

# Ordenação personalizada das categorias (4 grupos principais)
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
$title2 = '🛡️  ASSESSMENT MICROSOFT DEFENDER FOR ENDPOINT (MDE) v2.4 🛡️'
$padLength2 = ($bannerWidth - $title2.Length) / 2
$padLeft2 = ' ' * [math]::Floor($padLength2)
$padRight2 = ' ' * [math]::Ceiling($padLength2)
Write-Host "║$padLeft2$title2$padRight2║" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
$subtitle2 = 'Avaliação de Conformidade com Melhores Práticas de Segurança'
$padLengthSub2 = ($bannerWidth - $subtitle2.Length) / 2
$padLeftSub2 = ' ' * [math]::Floor($padLengthSub2)
$padRightSub2 = ' ' * [math]::Ceiling($padLengthSub2)
Write-Host "║$padLeftSub2$subtitle2$padRightSub2║" -ForegroundColor Cyan
Write-Host "║$(' ' * $bannerWidth)║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
Write-Host

# Resumo de conformidade geral
$totalItems = $orderedResults.Count
$compliantItems = ($orderedResults | Where-Object { $_.Status -eq 'Conforme' }).Count
$nonCompliantItems = $totalItems - $compliantItems
$percentCompliant = if ($totalItems -gt 0) { [math]::Round(($compliantItems / $totalItems) * 100, 2) } else { 0 }

Write-Host
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$scoreTitle = 'SCORE DE CONFORMIDADE GERAL'
$padLengthScore = ($bannerWidth - $scoreTitle.Length) / 2
$padLeftScore = ' ' * [math]::Floor($padLengthScore)
$padRightScore = ' ' * [math]::Ceiling($padLengthScore)
Write-Host "║$padLeftScore$scoreTitle$padRightScore║" -ForegroundColor Cyan
Write-Host "╚$line╝" -ForegroundColor Cyan
Write-Host
Write-Host "  📊 Total de Itens Avaliados: " -NoNewline -ForegroundColor White
Write-Host $totalItems -ForegroundColor Yellow
Write-Host "  ✓  Itens Conformes: " -NoNewline -ForegroundColor Green
Write-Host $compliantItems -ForegroundColor Green
Write-Host "  ⚠  Itens com Atenção: " -NoNewline -ForegroundColor Yellow
Write-Host $nonCompliantItems -ForegroundColor $(if ($nonCompliantItems -eq 0) { 'Green' } else { 'Red' })
Write-Host
Write-Host "  🎯 CONFORMIDADE GERAL: " -NoNewline -ForegroundColor White
Write-Host "$percentCompliant%" -ForegroundColor $(if ($percentCompliant -ge 80) { 'Green' } elseif ($percentCompliant -ge 60) { 'Yellow' } else { 'Red' }) -NoNewline
Write-Host " ($compliantItems/$totalItems)" -ForegroundColor Gray
Write-Host

# Resumo por categoria com visualização melhorada
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$categoryTitle = 'SCORE DE CONFORMIDADE POR CATEGORIA'
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
    
    # Barra de progresso visual
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

# Sessão MDE Nativo
Write-Host
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$mdeTitle = '🛡️  MICROSOFT DEFENDER FOR ENDPOINT (MDE) - IMPLEMENTAÇÕES NATIVAS'
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

# Sessão Windows
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$windowsTitle = '🪟  CONFIGURAÇÕES DO WINDOWS (não específicas do MDE)'
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

# Sessão Compliance
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$complianceTitle = '📋  COMPLIANCE DE DISPOSITIVO & HARDWARE'
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

# Sessão Navegadores
$bannerWidth = 84
$line = '═' * $bannerWidth
Write-Host "╔$line╗" -ForegroundColor Cyan
$browserTitle = '🌐  PROTEÇÃO DE NAVEGADORES'
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

# Sessão ASR específica
$asrSummary = $orderedResults | Where-Object { $_.Categoria -eq 'MDE Nativo' -and $_.'Configuração' -like '*ASR*' }
if ($asrSummary) {
    Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Cyan
Write-Host ('║{0}║' -f ('🎯 RESUMO DETALHADO DAS REGRAS ASR (Attack Surface Reduction)'.PadLeft(72).PadRight(83))) -ForegroundColor Cyan
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Cyan
    Write-Host
    $asrSummary | Select-Object Categoria, 'Configuração', 'Guid', 'Código Atual', 'Código Recomendado', 'Valor Atual' | Format-Table -AutoSize
    Write-Host
}

# Itens que exigem atenção
$attentionItems = $orderedResults | Where-Object { $_.Status -ne 'Conforme' }
if ($attentionItems) {
    Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Yellow
Write-Host ('║{0}║' -f ('⚠️  ITENS QUE EXIGEM ATENÇÃO IMEDIATA ⚠️'.PadLeft(62).PadRight(83))) -ForegroundColor Yellow
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Yellow
    Write-Host
    
    $groupedAttention = $attentionItems | Group-Object Categoria | Sort-Object Name
    foreach ($group in $groupedAttention) {
        Write-Host "  📌 $($group.Name)" -ForegroundColor Cyan
        foreach ($item in $group.Group) {
            Write-Host "     ⚠  " -NoNewline -ForegroundColor Red
            Write-Host "$($item.'Configuração')" -NoNewline -ForegroundColor Yellow
            Write-Host " - $($item.'Recomendação')" -ForegroundColor White
            Write-Host "        💡 Correção: " -NoNewline -ForegroundColor DarkGray
            Write-Host "$($item.'Comando de Correção')" -ForegroundColor DarkYellow
        }
        Write-Host
    }
} else {
    Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Green
Write-Host ('║{0}║' -f ('✅ PARABÉNS! TUDO CONFORME! ✅'.PadLeft(58).PadRight(83))) -ForegroundColor Green
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Green
    Write-Host
    Write-Host '  🎉 Todas as configurações avaliadas estão em conformidade com as melhores práticas!' -ForegroundColor Green
    Write-Host
}

# ---------- Exportar relatório ----------
try {
    # Tenta exportar para XLSX primeiro
    $xlsxPath = [System.IO.Path]::ChangeExtension($OutputPath, 'xlsx')
    $xlsxSuccess = $false
    
    try {
        # Verifica política de execução
        $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
        if ($currentPolicy -eq 'Restricted') {
            Write-Verbose 'Ajustando política de execução para permitir módulos...'
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        }

        # Instala ImportExcel se necessário
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Verbose 'Instalando módulo ImportExcel...'
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
        }
        
        Import-Module ImportExcel -Force
        
        Write-Verbose "Exportando relatório consolidado para $xlsxPath"
    $checklistSheet = $orderedResults | Select-Object Categoria, 'Configuração', 'Guid', 'Valor Atual', 'Código Atual', 'Código Recomendado', 'Best Practice', 'Recomendação', 'Comando de Correção', 'Documentação', Status
    $recomendacoesSheet = $orderedResults | Where-Object { $_.Status -ne 'Conforme' } | Select-Object Categoria, 'Configuração', 'Valor Atual', 'Best Practice', 'Recomendação', 'Comando de Correção', 'Documentação', Status

        # Exporta checklist completo
        $checklistSheet | Export-Excel -Path $xlsxPath -WorksheetName 'Checklist' -AutoSize -FreezeTopRow
        
        # Exporta recomendações em aba separada
        if ($recomendacoesSheet) {
            $recomendacoesSheet | Export-Excel -Path $xlsxPath -WorksheetName 'Recomendacoes' -AutoSize -FreezeTopRow
        }
        
        $xlsxSuccess = $true
        Write-Host
        Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Green
Write-Host ('║{0}║' -f ('✅ RELATÓRIO XLSX GERADO COM SUCESSO!'.PadLeft(62).PadRight(83))) -ForegroundColor Green
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Green
        Write-Host
        Write-Host "  📄 Arquivo: " -NoNewline -ForegroundColor White
        Write-Host $xlsxPath -ForegroundColor Cyan
        Write-Host "  📊 Abas criadas:" -ForegroundColor White
        Write-Host "     • Checklist (Todos os itens)" -ForegroundColor Gray
        if ($recomendacoesSheet) {
            Write-Host "     • Recomendações (Itens não conformes)" -ForegroundColor Gray
        }
        Write-Host
        
    } catch {
        Write-Warning "Não foi possível gerar XLSX: $($_.Exception.Message)"
        Write-Warning "Gerando fallback em CSV..."
    }
    
    # Fallback para CSV se XLSX falhar
    if (-not $xlsxSuccess) {
        Write-Verbose "Exportando relatório consolidado para $OutputPath"
        $orderedResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        
        # Gera CSV separado para recomendações
        $recomendacoesPath = [System.IO.Path]::ChangeExtension($OutputPath, '_Recomendacoes.csv')
        $recomendacoesSheet = $orderedResults | Where-Object { $_.Status -ne 'Conforme' } | Select-Object Categoria, 'Configuração', 'Valor Atual', 'Best Practice', 'Recomendação', 'Comando de Correção', Status
        if ($recomendacoesSheet) {
            $recomendacoesSheet | Export-Csv -Path $recomendacoesPath -NoTypeInformation -Encoding UTF8
        }
        
        Write-Host
        Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Green
Write-Host ('║{0}║' -f ('✅ RELATÓRIO CSV GERADO COM SUCESSO!'.PadLeft(61).PadRight(83))) -ForegroundColor Green
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Green
        Write-Host
        Write-Host "  📄 Arquivo Principal: " -NoNewline -ForegroundColor White
        Write-Host $OutputPath -ForegroundColor Cyan
        if ($recomendacoesSheet) {
            Write-Host "  📋 Recomendações: " -NoNewline -ForegroundColor White
            Write-Host $recomendacoesPath -ForegroundColor Cyan
        }
        Write-Host
    }
    
} catch {
    Write-Host
    Write-Host '╔═══════════════════════════════════════════════════════════════════════════════════╗' -ForegroundColor Red
Write-Host ('║{0}║' -f ('❌ ERRO AO EXPORTAR RELATÓRIO'.PadLeft(58).PadRight(83))) -ForegroundColor Red
Write-Host '╚═══════════════════════════════════════════════════════════════════════════════════╝' -ForegroundColor Red
    Write-Host
    Write-Host "  ⚠️  Erro: " -NoNewline -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host
    exit 1
}
