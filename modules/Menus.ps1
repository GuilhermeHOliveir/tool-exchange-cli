# SPDX-License-Identifier: MIT
# Copyright (c) 2026 GuilhermeHOliveir
function Show-RetentionMenu {
    while ($true) {
        Show-Header 'RETENCAO MRM'
        Show-MenuOption '1' 'Consultar politicas / contas'
        Show-MenuOption '2' 'Criar tags + politica'
        Show-MenuOption '3' 'Aplicar politica existente'
        $choice=Read-MenuValue 'Opcao'
        if ($null -eq $choice) { return }
        switch ($choice) {
            '1' { Ensure-ExchangeConnection; Show-Queries }
            '2' { Ensure-ExchangeConnection; Show-CreateMenu }
            '3' { Ensure-ExchangeConnection; Show-PolicyCatalog }
            default { Write-Host 'Opcao invalida.'; Wait-Menu }
        }
    }
}

function Show-ReportsMenu {
    while ($true) {
        Show-Header 'AUDITORIA / RELATORIOS (SOMENTE LEITURA)'
        Show-MenuOption '1' 'Usuarios sem licenca com mailbox diferente de SharedMailbox'
        Show-MenuOption '2' 'Usuarios licenciados / atividade de email D90'
        Show-MenuOption '3' 'Previa de SharedMailboxes sem licenca pendentes de .desativado'
        $choice=Read-MenuValue 'Opcao'
        if ($null -eq $choice) { return }
        if ($choice -notin @('1','2','3')) { Write-Host 'Opcao invalida.'; continue }
        Ensure-ExchangeConnection
        $scopes=@('User.Read.All')
        if ($choice -eq '2') { $scopes+='Reports.Read.All' }
        Ensure-GraphConnection $scopes
        switch ($choice) {
            '1' { Invoke-UnlicensedReport }
            '2' { Invoke-InactivityReport }
            '3' { Invoke-SharedPreview }
        }
        Wait-Menu
    }
}

function Select-ReviewSheet {
    while ($true) {
        Show-Header 'SELECIONAR PLANILHA REVISADA'
        $sheets=@($script:SessionReviewSheets | Where-Object { $_ -and (Test-Path -LiteralPath $_.Path -PathType Leaf) })
        for ($i=0; $i -lt $sheets.Count; $i++) {
            Show-MenuOption ([string]($i+1)) ("$($sheets[$i].Type) | $($sheets[$i].Count) registro(s) | $($sheets[$i].Path)")
        }
        $other=$sheets.Count+1
        Show-MenuOption ([string]$other) 'Informar caminho de outro CSV'
        $choice=Read-MenuValue 'Numero da planilha'
        if ($null -eq $choice) { return $null }
        $number=0
        if (-not [int]::TryParse($choice,[ref]$number)) { Write-Host 'Digite o numero da planilha.' -ForegroundColor Yellow; continue }
        if ($number -ge 1 -and $number -le $sheets.Count) { return $sheets[$number-1].Path }
        if ($number -eq $other) {
            $path=Read-RequiredValue 'Caminho do CSV validado'
            if ($null -eq $path) { continue }
            if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
            Write-Host 'Arquivo nao encontrado.' -ForegroundColor Yellow
            continue
        }
        Write-Host 'Opcao invalida.' -ForegroundColor Yellow
    }
}

function Show-MailboxesMenu {
    while ($true) {
        Show-Header 'CAIXAS / DESATIVACAO CONTROLADA'
        Show-MenuOption '1' 'Converter para SharedMailbox + .desativado (licencas opcionais)'
        Show-MenuOption '2' 'Aplicar .desativado em SharedMailboxes sem licenca'
        Write-MenuText 'Opcao 1: CSV revisado. Opcao 2: consulta direta, sem planilha.' DarkGray
        $choice=Read-MenuValue 'Opcao'
        if ($null -eq $choice) { return }
        if ($choice -notin @('1','2')) { Write-Host 'Opcao invalida.'; continue }
        if ($choice -eq '2') {
            Ensure-ExchangeConnection
            Ensure-GraphConnection @('User.Read.All')
            Invoke-SharedRenameInteractive
            return
        }
        $path=Select-ReviewSheet
        if ($null -eq $path) { continue }
        $licenseChoice=Read-MenuValue '[1] Preservar licencas | [2] Remover licencas diretas'
        if ($null -eq $licenseChoice) { continue }
        if ($licenseChoice -notin @('1','2')) { Write-Host 'Opcao invalida.'; continue }
        $remove=$licenseChoice -eq '2'
        $mode=Read-ExecutionMode
        if ($null -eq $mode) { continue }
        Ensure-ExchangeConnection
        $scopes=@('User.Read.All')
        if ($remove -and $mode -eq 'Aplicar') {
            $scopes+='LicenseAssignment.ReadWrite.All'
            Import-Module Microsoft.Graph.Users.Actions -Global -ErrorAction Stop
        }
        Ensure-GraphConnection $scopes
        if ($mode -eq 'Simular') {
            $simulation=Invoke-MailboxChanges $path $true $remove -PassThru
            Complete-MailboxSimulation $simulation
            return
        }
        Invoke-MailboxChanges $path $true $remove $true
        Wait-Menu
    }
}

function Show-GroupsMenu {
    while ($true) {
        Show-Header 'GRUPOS / LISTAS'
        Show-MenuOption '1' 'Importar membros de CSV/TXT para lista de distribuicao'
        Show-MenuOption '2' 'Remover caixas .desativado sem licenca de grupos/listas'
        $choice=Read-MenuValue 'Opcao'
        if ($null -eq $choice) { return }
        switch ($choice) {
            '1' {
                $group=Read-RequiredValue 'Alias ou email exato do grupo'
                if ($null -eq $group) { continue }
                $path=Read-RequiredValue 'Caminho do CSV ou TXT'
                if ($null -eq $path) { continue }
                $column='Email'
                if ([IO.Path]::GetExtension($path) -ieq '.csv') {
                    $column=Read-RequiredValue 'Coluna dos emails (Enter = Email)' 'Email'
                    if ($null -eq $column) { continue }
                }
                $ownerChoice=Read-MenuValue '[1] Verificacao normal de proprietario | [2] Usar bypass administrativo'
                if ($null -eq $ownerChoice) { continue }
                if ($ownerChoice -notin @('1','2')) { Write-Host 'Opcao invalida.'; continue }
                $mode=Read-ExecutionMode
                if ($null -eq $mode) { continue }
                Ensure-ExchangeConnection
                Invoke-MemberImport $group $path $column ($mode -eq 'Aplicar') ($ownerChoice -eq '2')
                Wait-Menu
            }
            '2' {
                $ownerChoice=Read-MenuValue '[1] Preservar owners | [2] Incluir owners, preservando o ultimo'
                if ($null -eq $ownerChoice) { continue }
                if ($ownerChoice -notin @('1','2')) { Write-Host 'Opcao invalida.'; continue }
                $mode=Read-ExecutionMode
                if ($null -eq $mode) { continue }
                Ensure-ExchangeConnection
                Ensure-GraphConnection
                Invoke-GroupCleanup ($mode -eq 'Aplicar') ($ownerChoice -eq '2')
                Wait-Menu
            }
            default { Write-Host 'Opcao invalida.'; Wait-Menu }
        }
    }
}

function Show-ConnectionHelp {
    Show-Header 'CONEXOES / DEPENDENCIAS'
    Write-Host "PowerShell: $($PSVersionTable.PSVersion) | Programa: $script:AppRoot"
    foreach ($name in @('ExchangeOnlineManagement','Microsoft.Graph.Authentication','Microsoft.Graph.Users','Microsoft.Graph.Reports','Microsoft.Graph.Users.Actions')) {
        $installed=@(Get-Module -ListAvailable -Name $name)
        Write-Host ('{0}: {1}' -f $name, $(if ($installed.Count) { ($installed.Version | Sort-Object -Descending | Select-Object -First 1) } else { 'nao instalado' }))
    }
    if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        Get-ConnectionInformation | Format-Table UserPrincipalName,Organization,TenantID,State -AutoSize | Out-Host
    }
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        Get-MgContext | Select-Object Account,TenantId,Scopes | Format-List | Out-Host
    }
    Write-Host 'Autenticacao solicitada ao selecionar uma rotina. Nenhum modulo e instalado automaticamente.'
    Write-Host 'Use uma janela dedicada. Sessoes existentes sao reaproveitadas e nao sao desconectadas ao sair.'
    Write-Host "Registro: $script:LogPath"
    Wait-Menu
}

function Show-MainMenu {
    while ($true) {
        try {
            Show-Header 'MENU PRINCIPAL'
            Show-MenuOption '1' 'Retencao MRM'
            Show-MenuOption '2' 'Auditoria / relatorios'
            Show-MenuOption '3' 'Caixas / desativacao controlada'
            Show-MenuOption '4' 'Grupos / listas'
            Show-MenuOption '5' 'Conexoes / dependencias'
            $choice=Read-MenuValue 'Opcao'
            if ($null -eq $choice) { return }
            switch ($choice) {
                '1' { Show-RetentionMenu }
                '2' { Show-ReportsMenu }
                '3' { Show-MailboxesMenu }
                '4' { Show-GroupsMenu }
                '5' { Show-ConnectionHelp }
                default { Write-Host 'Opcao invalida.'; Wait-Menu }
            }
        } catch {
            if ($_.Exception.Message -eq 'MRM_MENU_HOME') { continue }
            if ($_.Exception.Message -eq 'MRM_MENU_EXIT') { return }
            Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host 'Fluxo interrompido. Consulte o registro antes de repetir uma alteracao.' -ForegroundColor Yellow
            try { Wait-Menu } catch {
                if ($_.Exception.Message -eq 'MRM_MENU_EXIT') { return }
                if ($_.Exception.Message -ne 'MRM_MENU_HOME') { throw }
            }
        }
    }
}

function Start-ExchangeAdmin {
    $ErrorActionPreference='Stop'
    try { Initialize-Audit; $script:SessionReviewSheets=@(); Show-MainMenu }
    finally { Write-Host "Encerrado. Registro local: $script:LogPath" }
}
