# AD Admin Tool

Ferramenta gráfica PowerShell para gerenciamento de usuários no Active Directory.

## ✅ Pré-requisitos
- Windows 10/11 ou Server com RSAT
- PowerShell 5.1+
- Módulo ActiveDirectory instalado
- Permissões de administrador no AD

## ⚡ Funcionalidades
- Desbloquear contas AD
- Redefinir senhas
- Verificar status de usuários

## 📦 Compilar para EXE
Para converter o script em executável:

1. Instale o módulo ps2exe no powershell:

-Install-Module -Name ps2exe -Scope CurrentUser -Force

2. Compile o script.

-Invoke-ps2exe -inputFile "ADAdminTool.ps1" -outputFile "ADAdminTool.exe" -title "AD Admin Tool" -noConsole

Obs: Caso queira adicionar um icone personalizado ao executável, inclua o parâmetro -iconFile com o caminho do arquivo .ico
