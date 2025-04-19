Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

## Formulário Principal (05/08/2023)##
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Administracao de Usuario (AD-DC)"
$mainForm.Size = New-Object System.Drawing.Size(650, 600)  # Aumentado para acomodar mais informações
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false
$mainForm.BackColor = [System.Drawing.Color]::White

# Cabeçalho
$header = New-Object System.Windows.Forms.Label
$header.Text = "ADMINISTRACAO USUARIOS - Secretaria De Saude(CLPTA)-by Gustavo"
$header.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$header.Location = New-Object System.Drawing.Point(20, 15)
$header.Size = New-Object System.Drawing.Size(600, 25)
$header.ForeColor = [System.Drawing.Color]::DarkSlateBlue
$mainForm.Controls.Add($header)

# Campo de busca
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Buscar usuario (nome ou login):"
$searchLabel.Location = New-Object System.Drawing.Point(20, 60)
$searchLabel.Size = New-Object System.Drawing.Size(200, 20)
$searchLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$mainForm.Controls.Add($searchLabel)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(20, 85)
$searchBox.Size = New-Object System.Drawing.Size(400, 20)
$searchBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$mainForm.Controls.Add($searchBox)

$searchBtn = New-Object System.Windows.Forms.Button
$searchBtn.Text = "BUSCAR"
$searchBtn.Location = New-Object System.Drawing.Point(430, 85)
$searchBtn.Size = New-Object System.Drawing.Size(80, 23)
$searchBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$searchBtn.BackColor = [System.Drawing.Color]::LightGray
$mainForm.Controls.Add($searchBtn)

# Lista de usuários encontrados
$usersList = New-Object System.Windows.Forms.ListBox
$usersList.Location = New-Object System.Drawing.Point(20, 120)
$usersList.Size = New-Object System.Drawing.Size(400, 150)
$usersList.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$usersList.SelectionMode = "MultiExtended"
$mainForm.Controls.Add($usersList)

# Painel de informações expandido
$infoPanel = New-Object System.Windows.Forms.Panel
$infoPanel.Location = New-Object System.Drawing.Point(20, 280)
$infoPanel.Size = New-Object System.Drawing.Size(600, 220)  # Aumentado para mais informações
$infoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$infoPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
$mainForm.Controls.Add($infoPanel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 10)
$statusLabel.Size = New-Object System.Drawing.Size(580, 200)  # Aumentado
$statusLabel.Font = New-Object System.Drawing.Font("Consolas", 8.5)  # Fonte menor para mais conteúdo
$statusLabel.Text = "Busque um usuario e selecione na lista para ver detalhes"
$infoPanel.Controls.Add($statusLabel)

# Painel de ação
$actionPanel = New-Object System.Windows.Forms.Panel
$actionPanel.Location = New-Object System.Drawing.Point(20, 510)
$actionPanel.Size = New-Object System.Drawing.Size(600, 50)
$mainForm.Controls.Add($actionPanel)

$unlockBtn = New-Object System.Windows.Forms.Button
$unlockBtn.Text = "DESBLOQUEAR"
$unlockBtn.Location = New-Object System.Drawing.Point(0, 0)
$unlockBtn.Size = New-Object System.Drawing.Size(100, 30)
$unlockBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$unlockBtn.BackColor = [System.Drawing.Color]::LightSteelBlue
$unlockBtn.Enabled = $false
$actionPanel.Controls.Add($unlockBtn)

$passwordBtn = New-Object System.Windows.Forms.Button
$passwordBtn.Text = "ALTERAR SENHA"
$passwordBtn.Location = New-Object System.Drawing.Point(110, 0)
$passwordBtn.Size = New-Object System.Drawing.Size(120, 30)
$passwordBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$passwordBtn.BackColor = [System.Drawing.Color]::LightSteelBlue
$passwordBtn.Enabled = $false
$actionPanel.Controls.Add($passwordBtn)

$optionsBtn = New-Object System.Windows.Forms.Button
$optionsBtn.Text = "OPCOES"
$optionsBtn.Location = New-Object System.Drawing.Point(240, 0)
$optionsBtn.Size = New-Object System.Drawing.Size(100, 30)
$optionsBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$optionsBtn.BackColor = [System.Drawing.Color]::LightSteelBlue
$optionsBtn.Enabled = $false
$actionPanel.Controls.Add($optionsBtn)

# Formulário de Senha
$passForm = New-Object System.Windows.Forms.Form
$passForm.Text = "Alteracao de Senha"
$passForm.Size = New-Object System.Drawing.Size(350, 300)
$passForm.StartPosition = "CenterScreen"
$passForm.FormBorderStyle = "FixedDialog"
$passForm.MaximizeBox = $false
$passForm.BackColor = [System.Drawing.Color]::White

$passHeader = New-Object System.Windows.Forms.Label
$passHeader.Text = "ALTERAR SENHA"
$passHeader.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$passHeader.Location = New-Object System.Drawing.Point(20, 15)
$passHeader.Size = New-Object System.Drawing.Size(300, 25)
$passHeader.ForeColor = [System.Drawing.Color]::DarkSlateBlue
$passForm.Controls.Add($passHeader)

$newPassLabel = New-Object System.Windows.Forms.Label
$newPassLabel.Text = "Nova senha:"
$newPassLabel.Location = New-Object System.Drawing.Point(20, 60)
$newPassLabel.Size = New-Object System.Drawing.Size(150, 20)
$newPassLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$passForm.Controls.Add($newPassLabel)

$newPassBox = New-Object System.Windows.Forms.TextBox
$newPassBox.PasswordChar = '*'
$newPassBox.Location = New-Object System.Drawing.Point(20, 85)
$newPassBox.Size = New-Object System.Drawing.Size(300, 20)
$newPassBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$passForm.Controls.Add($newPassBox)

$confirmPassLabel = New-Object System.Windows.Forms.Label
$confirmPassLabel.Text = "Confirmar senha:"
$confirmPassLabel.Location = New-Object System.Drawing.Point(20, 115)
$confirmPassLabel.Size = New-Object System.Drawing.Size(150, 20)
$confirmPassLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$passForm.Controls.Add($confirmPassLabel)

$confirmPassBox = New-Object System.Windows.Forms.TextBox
$confirmPassBox.PasswordChar = '*'
$confirmPassBox.Location = New-Object System.Drawing.Point(20, 140)
$confirmPassBox.Size = New-Object System.Drawing.Size(300, 20)
$confirmPassBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$passForm.Controls.Add($confirmPassBox)

$changeOption = New-Object System.Windows.Forms.CheckBox
$changeOption.Text = "Usuario deve alterar senha no proximo logon"
$changeOption.Location = New-Object System.Drawing.Point(20, 170)
$changeOption.Size = New-Object System.Drawing.Size(300, 20)
$changeOption.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$passForm.Controls.Add($changeOption)

$submitPassBtn = New-Object System.Windows.Forms.Button
$submitPassBtn.Text = "CONFIRMAR"
$submitPassBtn.Location = New-Object System.Drawing.Point(100, 210)
$submitPassBtn.Size = New-Object System.Drawing.Size(120, 30)
$submitPassBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$submitPassBtn.BackColor = [System.Drawing.Color]::SteelBlue
$submitPassBtn.ForeColor = [System.Drawing.Color]::White
$passForm.Controls.Add($submitPassBtn)

# Formulário de Opções (atualizado com mais opções 17/04/2025)
$optionsForm = New-Object System.Windows.Forms.Form
$optionsForm.Text = "Opcoes do Usuario"
$optionsForm.Size = New-Object System.Drawing.Size(350, 350)
$optionsForm.StartPosition = "CenterScreen"
$optionsForm.FormBorderStyle = "FixedDialog"
$optionsForm.MaximizeBox = $false
$optionsForm.BackColor = [System.Drawing.Color]::White

$optionsHeader = New-Object System.Windows.Forms.Label
$optionsHeader.Text = "CONFIGURACOES DO USUARIO"
$optionsHeader.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$optionsHeader.Location = New-Object System.Drawing.Point(20, 15)
$optionsHeader.Size = New-Object System.Drawing.Size(300, 25)
$optionsHeader.ForeColor = [System.Drawing.Color]::DarkSlateBlue
$optionsForm.Controls.Add($optionsHeader)

$enableAccount = New-Object System.Windows.Forms.CheckBox
$enableAccount.Text = "Conta habilitada"
$enableAccount.Location = New-Object System.Drawing.Point(20, 60)
$enableAccount.Size = New-Object System.Drawing.Size(150, 20)
$enableAccount.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$optionsForm.Controls.Add($enableAccount)

$pwdNeverExpires = New-Object System.Windows.Forms.CheckBox
$pwdNeverExpires.Text = "Senha nunca expira"
$pwdNeverExpires.Location = New-Object System.Drawing.Point(20, 90)
$pwdNeverExpires.Size = New-Object System.Drawing.Size(150, 20)
$pwdNeverExpires.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$optionsForm.Controls.Add($pwdNeverExpires)

$cannotChangePwd = New-Object System.Windows.Forms.CheckBox
$cannotChangePwd.Text = "Usuario nao pode alterar senha"
$cannotChangePwd.Location = New-Object System.Drawing.Point(20, 120)
$cannotChangePwd.Size = New-Object System.Drawing.Size(200, 20)
$cannotChangePwd.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$optionsForm.Controls.Add($cannotChangePwd)

$expireAccount = New-Object System.Windows.Forms.CheckBox
$expireAccount.Text = "Definir data de expiracao"
$expireAccount.Location = New-Object System.Drawing.Point(20, 150)
$expireAccount.Size = New-Object System.Drawing.Size(180, 20)
$expireAccount.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$optionsForm.Controls.Add($expireAccount)

$expireDatePicker = New-Object System.Windows.Forms.DateTimePicker
$expireDatePicker.Location = New-Object System.Drawing.Point(20, 180)
$expireDatePicker.Size = New-Object System.Drawing.Size(200, 20)
$expireDatePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
$expireDatePicker.Enabled = $false
$optionsForm.Controls.Add($expireDatePicker)

$expireAccount.Add_CheckedChanged({
    $expireDatePicker.Enabled = $expireAccount.Checked
})

$submitOptionsBtn = New-Object System.Windows.Forms.Button
$submitOptionsBtn.Text = "SALVAR"
$submitOptionsBtn.Location = New-Object System.Drawing.Point(100, 220)
$submitOptionsBtn.Size = New-Object System.Drawing.Size(120, 30)
$submitOptionsBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$submitOptionsBtn.BackColor = [System.Drawing.Color]::SteelBlue
$submitOptionsBtn.ForeColor = [System.Drawing.Color]::White
$optionsForm.Controls.Add($submitOptionsBtn)

## FUNÇÕES PRINCIPAIS ATUALIZADAS ##
function Search-Users {
    $searchText = $searchBox.Text.Trim()
    
    if (-not $searchText) {
        [System.Windows.Forms.MessageBox]::Show("Digite um nome ou login para buscar", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    try {
        $users = Get-ADUser -Filter "Name -like '*$searchText*' -or SamAccountName -like '*$searchText*'" -Properties Name, SamAccountName, Enabled, EmailAddress
        
        $usersList.Items.Clear()
        foreach ($user in $users) {
            $status = if ($user.Enabled) { "(Ativo)" } else { "(Inativo)" }
            $email = if ($user.EmailAddress) { "| $($user.EmailAddress)" } else { "" }
            $usersList.Items.Add("$($user.Name) [$($user.SamAccountName)] $status $email")
        }
        
        if ($users.Count -eq 0) {
            $statusLabel.Text = "Nenhum usuario encontrado com '$searchText'"
            $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
        }
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Erro na busca: $_", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Show-UserDetails {
    if ($usersList.SelectedIndex -eq -1) {
        return
    }

    $selectedText = $usersList.SelectedItem.ToString()
    $samAccountName = ($selectedText -split '\[|\]')[1].Trim()
    
    try {
        $user = Get-ADUser -Identity $samAccountName -Properties *
        
        $infoText = "INFORMACOES BASICAS:`n"
        $infoText += "----------------------------------------`n"
        $infoText += "NOME COMPLETO: $($user.Name)`n"
        $infoText += "LOGIN: $($user.SamAccountName)`n"
        $infoText += "EMAIL: $($user.EmailAddress)`n"
        $infoText += "TELEFONE: $($user.OfficePhone)`n"
        $infoText += "DEPARTAMENTO: $($user.Department)`n"
        $infoText += "CARGO: $($user.Title)`n"
        $infoText += "LOCALIZACAO: $($user.PhysicalDeliveryOfficeName)`n"
        $infoText += "DESCRICAO: $($user.Description)`n"
        $infoText += "GERENTE: $(if($user.Manager) {(Get-ADUser $user.Manager -Properties Name).Name} else {'NAO DEFINIDO'})`n"
        
        $infoText += "`nSTATUS DA CONTA:`n"
        $infoText += "----------------------------------------`n"
        $infoText += "CONTA ATIVA: " + $(if ($user.Enabled) { "SIM" } else { "NAO" }) + "`n"
        $infoText += "CONTA BLOQUEADA: " + $(if ($user.LockedOut) { "SIM" } else { "NAO" }) + "`n"
        $infoText += "DATA DE CRIACAO: $($user.Created.ToString('dd/MM/yyyy'))`n"
        $infoText += "DATA DE EXPIRACAO: $(if($user.AccountExpirationDate) {$user.AccountExpirationDate.ToString('dd/MM/yyyy')} else {'NUNCA'})`n"
        
        $infoText += "`nINFORMACOES DE LOGON:`n"
        $infoText += "----------------------------------------`n"
        $infoText += "ULTIMO LOGON: " + $(if ($user.LastLogonDate) { $user.LastLogonDate.ToString("dd/MM/yyyy HH:mm") } else { "NUNCA" }) + "`n"
        $infoText += "HORARIO PERMITIDO: $(if($user.LogonHours) {'RESTRITO'} else {'SEM RESTRICAO'})`n"
        $infoText += "ESTACOES PERMITIDAS: $(if($user.LogonWorkstations) {'RESTRITO'} else {'TODAS'})`n"
        
        $infoText += "`nINFORMACOES DE SENHA:`n"
        $infoText += "----------------------------------------`n"
        $infoText += "SENHA ALTERADA: " + $user.PasswordLastSet.ToString("dd/MM/yyyy HH:mm") + "`n"
        $pwdAge = (New-TimeSpan -Start $user.PasswordLastSet -End (Get-Date)).Days
        $infoText += "IDADE DA SENHA: $pwdAge dias`n"
        $infoText += "DIAS PARA EXPIRAR: $(90 - $pwdAge) dias`n"
        $infoText += "PROXIMA TROCA: " + (($user.PasswordLastSet).AddDays(90).ToString("dd/MM/yyyy")) + "`n"
        $infoText += "ULTIMA TENTATIVA FALHA: " + $(if($user.LastBadPasswordAttempt) {$user.LastBadPasswordAttempt.ToString("dd/MM/yyyy HH:mm")} else {"NENHUMA"}) + "`n"
        $infoText += "TENTATIVAS FALHAS: $($user.BadPwdCount)`n"
        $infoText += "SENHA NUNCA EXPIRA: $(if ($user.PasswordNeverExpires) { "SIM" } else { "NAO" })`n"
        $infoText += "PODE ALTERAR SENHA: $(if ($user.CannotChangePassword) { "NAO" } else { "SIM" })`n"
        
        $infoText += "`nMEMBRESIA DE GRUPOS:`n"
        $infoText += "----------------------------------------`n"
        $groups = Get-ADPrincipalGroupMembership $user | Select-Object -First 8 -ExpandProperty Name
        $infoText += "GRUPOS PRINCIPAIS: $($groups -join ', ')`n"
        
        $statusLabel.Text = $infoText
        $statusLabel.ForeColor = [System.Drawing.Color]::Black
        
        $unlockBtn.Enabled = $user.LockedOut
        $passwordBtn.Enabled = $true
        $optionsBtn.Enabled = $true
        
    } catch {
        $statusLabel.Text = "Erro ao carregar detalhes do usuario: $_"
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkRed
    }
}

function Unlock-UserAccount {
    $selectedText = $usersList.SelectedItem.ToString()
    $samAccountName = ($selectedText -split '\[|\]')[1].Trim()
    
    try {
        Unlock-ADAccount -Identity $samAccountName
        $statusLabel.Text = "Usuario $samAccountName desbloqueado com sucesso!`n`n" + $statusLabel.Text
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        $unlockBtn.Enabled = $false
        
        # Atualiza os dados do usuario
        Show-UserDetails
        
    } catch {
        $statusLabel.Text = "Falha ao desbloquear usuario: $_`n`n" + $statusLabel.Text
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkRed
    }
}

function Change-UserPassword {
    $selectedText = $usersList.SelectedItem.ToString()
    $global:selectedUser = ($selectedText -split '\[|\]')[1].Trim()
    
    $passForm.ShowDialog()
}

function Save-NewPassword {
    if ($newPassBox.Text -ne $confirmPassBox.Text) {
        [System.Windows.Forms.MessageBox]::Show("As senhas nao coincidem!", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    try {
        $securePwd = ConvertTo-SecureString $newPassBox.Text -AsPlainText -Force
        Set-ADAccountPassword -Identity $global:selectedUser -NewPassword $securePwd -Reset
        Set-ADUser -Identity $global:selectedUser -ChangePasswordAtLogon $changeOption.Checked
        
        $passForm.Close()
        $statusLabel.Text = "Senha alterada com sucesso para $global:selectedUser`n`n" + $statusLabel.Text
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        
        $newPassBox.Text = ""
        $confirmPassBox.Text = ""
        $changeOption.Checked = $false
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Falha ao alterar senha: $_", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Show-UserOptions {
    $selectedText = $usersList.SelectedItem.ToString()
    $global:selectedUser = ($selectedText -split '\[|\]')[1].Trim()
    
    try {
        $user = Get-ADUser -Identity $global:selectedUser -Properties Enabled, PasswordNeverExpires, CannotChangePassword, AccountExpirationDate
        
        $enableAccount.Checked = $user.Enabled
        $pwdNeverExpires.Checked = $user.PasswordNeverExpires
        $cannotChangePwd.Checked = $user.CannotChangePassword
        
        if ($user.AccountExpirationDate) {
            $expireAccount.Checked = $true
            $expireDatePicker.Value = $user.AccountExpirationDate
        } else {
            $expireAccount.Checked = $false
            $expireDatePicker.Value = (Get-Date).AddYears(1)
        }
        
        $optionsForm.ShowDialog()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Falha ao carregar opcoes: $_", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Save-UserOptions {
    try {
        $params = @{
            Identity = $global:selectedUser
            Enabled = $enableAccount.Checked
            PasswordNeverExpires = $pwdNeverExpires.Checked
            CannotChangePassword = $cannotChangePwd.Checked
        }
        
        if ($expireAccount.Checked) {
            $params['AccountExpirationDate'] = $expireDatePicker.Value
        } else {
            $params['AccountExpirationDate'] = $null
        }
        
        Set-ADUser @params
        
        $optionsForm.Close()
        $statusLabel.Text = "Opcoes atualizadas para $global:selectedUser`n`n" + $statusLabel.Text
        $statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        
        # Atualiza a lista para refletir mudanças
        Search-Users
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Falha ao salvar opcoes: $_", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

## EVENTOS ##
$searchBtn.Add_Click({ Search-Users })
$usersList.Add_SelectedIndexChanged({ Show-UserDetails })
$unlockBtn.Add_Click({ Unlock-UserAccount })
$passwordBtn.Add_Click({ Change-UserPassword })
$optionsBtn.Add_Click({ Show-UserOptions })
$submitPassBtn.Add_Click({ Save-NewPassword })
$submitOptionsBtn.Add_Click({ Save-UserOptions })

# Enter para buscar
$searchBox.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Search-Users
    }
})

# Iniciar aplicação
$mainForm.Add_Shown({ $searchBox.Focus() })
[void] $mainForm.ShowDialog()
