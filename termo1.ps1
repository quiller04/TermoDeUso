# Carregar o arquivo de credenciais
. "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\termo1\credentials.ps1"

# Função para enviar e-mail usando SMTP com conta Microsoft
function EnviarEmail {
    param (
        [string]$de,
        [string]$para,
        [string]$assunto,
        [string]$corpo
    )

    try {
        # Configurar credenciais do remetente
        $senhaSegura = ConvertTo-SecureString $senhaApp -AsPlainText -Force
        $credenciaisEmail = New-Object System.Management.Automation.PSCredential($de, $senhaSegura)

        # Configurar o e-mail com codificação UTF-8
        $mailMessage = New-Object system.net.mail.mailmessage
        $mailMessage.From = $de
        $mailMessage.To.Add($para)
        $mailMessage.Subject = $assunto
        $mailMessage.Body = $corpo
        $mailMessage.IsBodyHtml = $true
        $mailMessage.BodyEncoding = [System.Text.Encoding]::UTF8

        # Configurar o cliente SMTP
        $smtpClient = New-Object system.net.mail.smtpclient("smtp.office365.com", 587)
        $smtpClient.EnableSsl = $true
        $smtpClient.Credentials = $credenciaisEmail

        # Enviar o e-mail
        $smtpClient.Send($mailMessage)
        Write-Host "E-mail enviado com sucesso para $para"
    }
    catch {
        Write-Host "Erro ao enviar e-mail: $_"
    }
}

# Função para obter o Serial Number e o Modelo da máquina
function Get-ComputerInfo {
    $computerInfo = Get-WmiObject -Class Win32_BIOS
    $computerModel = Get-WmiObject -Class Win32_ComputerSystem

    return @{
        SerialNumber = $computerInfo.SerialNumber
        Model        = $computerModel.Model
    }
}

# Função para ler um arquivo .txt e substituir placeholders
function PreencherCampos {
    param (
        [string]$filePath,
        [hashtable]$substituicoes
    )

    # Carregar o conteúdo do arquivo com codificação UTF-8
    $conteudo = Get-Content $filePath -Raw -Encoding UTF8

    # Substituir placeholders e adicionar quebras de linha HTML
    foreach ($chave in $substituicoes.Keys) {
        $conteudo = $conteudo -replace $chave, $substituicoes[$chave]
    }
    $conteudo = $conteudo -replace "`r`n", "<br>"

    return $conteudo
}

# Carregar o módulo do Active Directory
Import-Module ActiveDirectory

# Configurar as credenciais de administração para modificar o AD
$usuarioAdmin = "arezzo.local\admin.gaalencar"
$credenciais = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $usuarioAdmin, (ConvertTo-SecureString $senhaAdminAD -AsPlainText -Force)

# Pegar o nome do usuário logado e o hostname da máquina
$usuario = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[1]
$hostname = $env:COMPUTERNAME

# Buscar o usuário no AD e obter o atributo 'groupPriority' e 'userPrincipalName'
$ADUser = Get-ADUser -Identity $usuario -Properties groupPriority, userPrincipalName

# Verificar se o hostname atual já está presente no atributo 'groupPriority'
$hostnamesRegistrados = if ($ADUser.groupPriority) { $ADUser.groupPriority -split ";" } else { @() }

# Função para exibir o formulário do termo de uso
function ExibirTermoDeUso {
    # Carregar Assemblies para criar a UI
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    # Criar o formulário de aceite do termo de uso
    $form = New-Object Windows.Forms.Form
    $form.Text = "Termo de Uso"
    $form.Size = New-Object Drawing.Size(600, 500)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ControlBox = $false
    $form.TopMost = $true

    # Configuração do campo de texto do termo
    $textoTermo = New-Object Windows.Forms.RichTextBox
    $textoTermo.Size = New-Object Drawing.Size(580, 350)
    $textoTermo.Location = New-Object Drawing.Point(10, 10)
    $textoTermo.ReadOnly = $true
    $textoTermo.Multiline = $true
    $textoTermo.ScrollBars = 'Vertical'
    $textoTermo.Font = New-Object Drawing.Font("Arial", 10)
    $textoTermo.Text = [System.IO.File]::ReadAllText("$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\termo1\termo.txt", [System.Text.Encoding]::UTF8)
    $form.Controls.Add($textoTermo)

    # Configuração do checkbox de aceite
    $checkBox = New-Object Windows.Forms.CheckBox
    $checkBox.Size = New-Object Drawing.Size(200, 20)
    $checkBox.Location = New-Object Drawing.Point(10, 370)
    $checkBox.Text = "Li o termo e estou ciente"
    $form.Controls.Add($checkBox)

    # Configuração do botão de aceite
    $botaoAceitar = New-Object Windows.Forms.Button
    $botaoAceitar.Size = New-Object Drawing.Size(100, 30)
    $botaoAceitar.Location = New-Object Drawing.Point(450, 420)
    $botaoAceitar.Text = "Aceito"
    $botaoAceitar.Enabled = $false
    $form.Controls.Add($botaoAceitar)

    # Habilitar o botão quando o checkbox for marcado
    $checkBox.Add_CheckedChanged({
            $botaoAceitar.Enabled = $checkBox.Checked
        })

    # Ação do botão "Aceito"
    $botaoAceitar.Add_Click({
            # Minimizar o formulário e atualizar o campo 'groupPriority' no AD
            $form.WindowState = 'Minimized'
            if ($ADUser.groupPriority) {
                "$($ADUser.groupPriority);$hostname"
            }
            else {
                $hostname
            }

            Set-ADUser -Identity $usuario -Add @{groupPriority = $hostname } -Credential $credenciais
            Write-Host "Termo aceito e hostname $hostname adicionado ao atributo 'groupPriority'."

            # Preparar informações e substituir placeholders nos e-mails
            $machineInfo = Get-ComputerInfo
            $substituicoes = @{
                '{usuario}'      = $usuario
                '{hostname}'     = $hostname
                '{serialNumber}' = $machineInfo.SerialNumber
                '{model}'        = $machineInfo.Model
                '{emailUsuario}' = $ADUser.userPrincipalName
            }

            # Preencher e enviar e-mail de confirmação ao usuário
            $conteudoEmailUsuario = PreencherCampos -filePath "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\termo1\email_usuario.txt" -substituicoes $substituicoes
            EnviarEmail -de "termodeuso@azzas2154.com.br" -para $ADUser.userPrincipalName -assunto "Aceite do Termo de Uso" -corpo $conteudoEmailUsuario

            # Preencher e enviar e-mail ao suporte
            $conteudoEmailSuporte = PreencherCampos -filePath "$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\termo1\email_suporte.txt" -substituicoes $substituicoes
            EnviarEmail -de "termodeuso@azzas2154.com.br" -para "termotisp@arezzo.com.br" -assunto "Aceite de Termo de Uso - $usuario" -corpo $conteudoEmailSuporte

            # Fechar o formulário após envio dos e-mails
            $form.Close()
        })

    # Adicionar imagem abaixo do checkbox
    $pictureBox = New-Object Windows.Forms.PictureBox
    $pictureBox.Size = New-Object Drawing.Size(300, 50)
    $pictureBox.Location = New-Object Drawing.Point(10, 400)
    $pictureBox.SizeMode = 'Zoom'
    $pictureBox.Image = [System.Drawing.Image]::FromFile("$env:USERPROFILE\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\termo1\azzas.png")
    $form.Controls.Add($pictureBox)

    # Exibir o formulário
    $form.ShowDialog()
}

# Executar a exibição do termo de uso se o hostname ainda não estiver registrado
if ($usuario -notlike "admin*") {
    if ($hostnamesRegistrados -notcontains $hostname) {
        ExibirTermoDeUso
    }
    else {
        Write-Host "Hostname $hostname já está registrado para o usuário $usuario."
    }
}
else {
    Write-Host "Execução bloqueada para o usuário admin."
}
