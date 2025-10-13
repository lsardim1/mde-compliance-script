# 🛡️ MDE Security Assessment Tool v2.4

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)](https://docs.microsoft.com/en-us/powershell/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Windows](https://img.shields.io/badge/Windows-10%2F11-blue)](https://www.microsoft.com/windows)

> **Complete Microsoft Defender for Endpoint compliance assessment**

Security audit script that evaluates endpoint configurations against Microsoft best practices, Intune baseline and advanced Windows controls, with enhanced visual interface and detailed reports.

## Features

### **MDE Native** - Defender for Endpoint Native Implementations
- **MDE/EDR Onboarding** (Sense service)
- **Microsoft Defender Antivirus**:
  - Real-time protection, cloud-delivered protection, MAPS
  - PUA (Potentially Unwanted Applications)
  - Network Protection, Behavior Monitoring
  - Block at First Sight (BAFS)
  - Tamper Protection
- **Attack Surface Reduction (ASR)** - 19 official rules with GUID and action code
- **Scan configurations** (scheduling, parameters)
- **Signature updates** and performance

### **Windows** - System Settings (non-MDE specific)
- **Encryption**: BitLocker (OS and removable media)
- **Windows Firewall** (profiles and auditing)
- **Windows Defender SmartScreen** (system-level protection)
- **Credential Guard** (credential protection)
- **Secure Boot** and **TPM 2.0**
- **WDAC** (Windows Defender Application Control)
- **Exploit Protection** (exploit mitigation)
- **Application Guard** (browser isolation)
- **UAC** (User Account Control)
- **LAPS** (Local Administrator Password Solution)
- Account policies (password, lockout)
- Auditing and logs
- Windows Update

### **Compliance** - Device & Hardware Compliance
- Device policies: minimum password length, screen lock
- **Windows Hello for Business** (biometric authentication)
- **TPM driver** updated

### **Browsers** - Browser Protection
- **Microsoft Edge SmartScreen**
- **Google Chrome Safe Browsing**
- **Mozilla Firefox Enhanced Tracking Protection**

## Installation and Usage

### Prerequisites
- **PowerShell 5.1** or higher
- **Run as Administrator** (recommended for full access)
- **Windows 10/11** with Microsoft Defender enabled
- **Internet connection** (for automatic ImportExcel module installation, if needed)
- **Execution policy**: automatically adjusted if necessary

### Usage Commands

#### 1. **Basic Execution** (Recommended)
```powershell
# Open PowerShell as Administrator and run:
.\Assessment-MDE-V2.4.ps1
```
> Generates report in `C:\temp\MDE_Assessment_Report_YYYY-MM-DD_HH-MM.xlsx`

#### 2. **Custom Path**
```powershell
.\Assessment-MDE-V2.4.ps1 -OutputPath "C:\Reports\Assessment.xlsx"
```

#### 3. **Verbose Mode** (Progress Details)
```powershell
.\Assessment-MDE-V2.4.ps1 -Verbose
```

#### 4. **Adjust Execution Policy** (if needed)
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Version 2.4 Features

### **New in v2.4:**
- **Complete reorganization**: clear separation between MDE Native and Windows
- **Firefox protection** (Enhanced Tracking Protection)
- **Visual interface** with boxes, icons and progress bars
- **Compliance score** with graphical visualization by category
- **Real-time progress indicators** during execution
- **Smart grouping** of non-compliant items by category
- **Validated documentation links** (34 URLs updated)
- **Enhanced export** with visual feedback (XLSX or CSV)
- **Error handling** with solution suggestions

### **Script Structure**
- **Monolithic**: Single file, no external dependencies
- **Utility functions**: Labels, conversions and formatting
- **Intune Baseline**: Microsoft recommended values integrated
- **Visual interface**: Colorful banners and clear categorization
- **Structured export**: XLSX/CSV ready for analysis

### **Performance and Characteristics**
- **Fast execution**: ~30-60 seconds
- **46+ security controls** evaluated
- **Enterprise focus**: Based on Microsoft best practices
- **Complete detail**: Each control with status, current value and recommendation
- **Compatibility**: Windows 10/11 and Windows Server

### Generated Reports

#### **1. Excel (XLSX) - 2 Tabs:**
- **Complete Checklist**: All controls with details
- **Recommendations**: Only items requiring attention

#### **2. CSV (Fallback)**: If ImportExcel module is not available

**Report Columns:**
| Field | Description |
|-------|-------------|
| **Category** | MDE Native / Windows / Compliance / Browsers |
| **Configuration** | Name of the evaluated control |
| **GUID** | Unique identifier (for ASR rules) |
| **Current Value** | State found in the system |
| **Best Practice** | Microsoft recommended value |
| **Status** | Compliant / Attention |
| **Remediation Command** | PowerShell/GPO for correction |
| **Documentation** | Official Microsoft link |

## Troubleshooting

### **"Execution Policy" Error**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### **"Access Denied" Error**
- Make sure to run as **Administrator**
- Verify that Windows Defender is active
- Confirm that user has local permissions

### **No WMI/CIM permissions**
```powershell
# Run as Administrator and check services:
Get-Service Winmgmt
Restart-Service Winmgmt -Force
```

### **ImportExcel module not installed**
The script tries to install automatically. If it fails:
```powershell
Install-Module -Name ImportExcel -Force -AllowClobber
```

### **Data collection failures**
- Wait a few seconds between executions
- Run `Get-MpPreference` manually to verify access
- Restart Windows Defender service if needed

## Documentation and Links

### **Official Microsoft Links:**
- [MDE Documentation](https://learn.microsoft.com/en-us/defender-endpoint/)
- [ASR Rules](https://learn.microsoft.com/en-us/defender-endpoint/attack-surface-reduction)
- [Intune Baselines](https://learn.microsoft.com/en-us/mem/intune/protect/security-baselines)
- [Windows Security](https://learn.microsoft.com/en-us/windows/security/)

### **Version History:**
- **v2.4** (10/07/2025): Category reorganization + Firefox + Visual interface
- **v2.2**: Control expansion + documentation + column separation
- **v2.0**: Complete baseline with 40+ controls
- **v1.0**: Initial version

## Author

**Leandro Sardim**
- **Last Update**: 10/07/2025
- **Contact**: Via GitHub Issues

## License

MIT License - Free use for commercial and educational purposes

## Contributions

For suggestions, improvements or bugs:
1. Open an [Issue](../../issues)
2. Fork the project  
3. Send a Pull Request

---

**If this project was helpful, leave a star!**

### **Quick Download**
```bash
# Repository clone
git clone https://github.com/YOUR_USER/mde-assessment-powershell.git
cd mde-assessment-powershell

# Direct execution
.\Assessment-MDE-V2.4.ps1
```

