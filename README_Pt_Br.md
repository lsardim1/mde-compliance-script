# 🛡️ MDE Security Assessment Tool v2.4

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)](https://docs.microsoft.com/en-us/powershell/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Windows](https://img.shields.io/badge/Windows-10%2F11-blue)](https://www.microsoft.com/windows)

> **Avaliação completa de conformidade do Microsoft Defender for Endpoint**

Script de auditoria de segurança que avalia configurações do endpoint contra melhores práticas Microsoft, baseline Intune e controles avançados do Windows, com interface visual aprimorada e relatórios detalhados.

## Funcionalidades

### **MDE Nativo** - Implementações Nativas do Defender for Endpoint
- **Onboarding MDE/EDR** (serviço Sense)
- **Microsoft Defender Antivírus**:
  - Proteção em tempo real, entregue pela nuvem, MAPS
  - PUA (Potentially Unwanted Applications)
  - Network Protection, Behavior Monitoring
  - Block at First Sight (BAFS)
  - Tamper Protection
- **Attack Surface Reduction (ASR)** - 19 regras oficiais com GUID e código de ação
- **Configurações de scan** (agendamento, parâmetros)
- **Atualização de assinaturas** e performance

### **Windows** - Configurações de Sistema (não específicas do MDE)
- **Criptografia**: BitLocker (OS e mídias removíveis)
- **Firewall** do Windows (perfis e auditoria)
- **Windows Defender SmartScreen** (proteção de sistema - nível OS)
- **Credential Guard** (proteção de credenciais)
- **Secure Boot** e **TPM 2.0**
- **WDAC** (Windows Defender Application Control)
- **Exploit Protection** (mitigação de exploits)
- **Application Guard** (isolamento de navegador)
- **UAC** (User Account Control)
- **LAPS** (Local Administrator Password Solution)
- Políticas de conta (senha, lockout)
- Auditoria e logs
- Windows Update

### **Compliance** - Conformidade de Dispositivo & Hardware
- Políticas de dispositivo: comprimento mínimo de senha, bloqueio de tela
- **Windows Hello for Business** (autenticação biométrica)
- **TPM driver** atualizado

### **Navegadores** - Proteção de Navegadores
- **Microsoft Edge SmartScreen**
- **Google Chrome Safe Browsing**
- **Mozilla Firefox Enhanced Tracking Protection**

## Instalação e Uso

### Pré-requisitos
- **PowerShell 5.1** ou superior
- **Execução como Administrador** (recomendado para acesso completo)
- **Windows 10/11** com Microsoft Defender habilitado
- **Conexão com internet** (para instalação automática do módulo ImportExcel, se necessário)
- **Política de execução**: ajustada automaticamente se necessário

### Comandos de Uso

#### 1. **Execução Básica** (Recomendada)
```powershell
# Abrir PowerShell como Administrador e executar:
.\Assessment-MDE-V2.4.ps1
```
> Gera relatório em `C:\temp\MDE_Assessment_Report_YYYY-MM-DD_HH-MM.xlsx`

#### 2. **Caminho Personalizado**
```powershell
.\Assessment-MDE-V2.4.ps1 -OutputPath "C:\Relatorios\Assessment.xlsx"
```

#### 3. **Modo Verbose** (Detalhes de Progresso)
```powershell
.\Assessment-MDE-V2.4.ps1 -Verbose
```

#### 4. **Ajustar Política de Execução** (se necessário)
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Recursos da Versão 2.4

### **Novidades v2.4:**
- **Reorganização completa**: separação clara entre MDE Nativo e Windows
- **Proteção para Firefox** (Enhanced Tracking Protection)
- **Interface visual profissional** com boxes, ícones e barras de progresso
- **Score de conformidade** com visualização gráfica por categoria
- **Indicadores de progresso** em tempo real durante execução
- **Agrupamento inteligente** de itens não conformes por categoria
- **Links de documentação** validados e atualizados (34 URLs)
- **Exportação aprimorada** com feedback visual (XLSX ou CSV)
- **Tratamento de erros** com sugestões de solução

### **Estrutura do Script**
- **Monolítico**: Arquivo único, sem dependências externas
- **Funções utilitárias**: Labels, conversões e formatação
- **Baseline Intune**: Valores recomendados Microsoft integrados
- **Interface visual**: Banners coloridos e categorização clara
- **Exportação estruturada**: XLSX/CSV pronto para análise

### **Performance e Características**
- **Execução rápida**: ~30-60 segundos
- **46+ controles** de segurança avaliados
- **Foco empresarial**: Baseado em melhores práticas Microsoft
- **Detalhamento completo**: Cada controle com status, valor atual e recomendação
- **Compatibilidade**: Windows 10/11 e Windows Server

### 📋 Relatórios Gerados

#### **1. Excel (XLSX) - 2 Abas:**
- **Checklist Completo**: Todos os controles com detalhes
- **Recomendações**: Apenas itens que requerem atenção

#### **2. CSV (Fallback)**: Se módulo ImportExcel não disponível

**Colunas do Relatório:**
| Campo | Descrição |
|-------|-----------|
| **Categoria** | MDE Nativo / Windows / Compliance / Navegadores |
| **Configuração** | Nome do controle avaliado |
| **GUID** | Identificador único (para regras ASR) |
| **Valor Atual** | Estado encontrado no sistema |
| **Best Practice** | Valor recomendado Microsoft |
| **Status** | Conforme / Atenção |
| **Comando de Correção** | PowerShell/GPO para correção |
| **Documentação** | Link oficial Microsoft |

## Troubleshooting

### **Erro "Execution Policy"**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

###  **Erro "Access Denied"**
-  Certifique-se de executar como **Administrador**
-  Verifique se o Windows Defender está ativo
-  Confirme que o usuário tem permissões locais

### **Sem permissões WMI/CIM**
```powershell
# Execute como Administrador e verifique serviços:
Get-Service Winmgmt
Restart-Service Winmgmt -Force
```

### **Módulo ImportExcel não instalado**
O script tenta instalar automaticamente. Se falhar:
```powershell
Install-Module -Name ImportExcel -Force -AllowClobber
```

###  **Falhas na coleta de dados**
-  Aguarde alguns segundos entre execuções
-  Execute `Get-MpPreference` manualmente para verificar acesso
-  Reinicie o serviço Windows Defender se necessário

## Documentação e Links

### **Links Oficiais Microsoft:**
- [MDE Documentação](https://learn.microsoft.com/en-us/defender-endpoint/)
- [ASR Rules](https://learn.microsoft.com/en-us/defender-endpoint/attack-surface-reduction)
- [Intune Baselines](https://learn.microsoft.com/en-us/mem/intune/protect/security-baselines)
- [Windows Security](https://learn.microsoft.com/en-us/windows/security/)

### **Histórico de Versões:**
- **v2.4** (07/10/2025): Reorganização categorias + Firefox + Interface visual profissional
- **v2.2**: Expansão de controles + documentação + separação de colunas
- **v2.0**: Baseline completo com 40+ controles
- **v1.0**: Versão inicial

##  Autor

**Leandro Sardim**
- **Última Atualização**: 07/10/2025
- **Contato**: Via GitHub Issues

## Licença

MIT License - Uso livre para fins comerciais e educacionais

## Contribuições

Para sugestões, melhorias ou bugs:
1.  Abra uma [Issue](../../issues)
2.  Faça um Fork do projeto  
3.  Envie um Pull Request

---

### **Download Rápido**
```bash
# Clone do repositório
git clone https://github.com/SEU_USUARIO/mde-assessment-powershell.git
cd mde-assessment-powershell

# Execução direta
.\Assessment-MDE-V2.4.ps1

```
