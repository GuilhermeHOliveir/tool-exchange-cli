# SPDX-License-Identifier: MIT
# Copyright (c) 2026 GuilhermeHOliveir
# Shared infrastructure. Importing the application never connects or changes a tenant.
function Read-RequiredValue([string]$Prompt, [string]$Default) {
    while ($true) {
        $value = Read-MenuValue $Prompt
        if ($null -eq $value) { return $null }
        if (-not $value -and $Default) { return $Default }
        if ($value) { return $value.Trim('"') }
        Write-Host 'Informe um valor.' -ForegroundColor Yellow
    }
}

function Read-ExecutionMode {
    while ($true) {
        $answer = Read-MenuValue '[1] Simular | [2] Aplicar em producao'
        if ($null -eq $answer) { return $null }
        if ($answer -eq '1') { return 'Simular' }
        if ($answer -eq '2') { return 'Aplicar' }
        Write-Host 'Opcao invalida.'
    }
}

function Confirm-Operation([string]$Message) {
    while ($true) {
        $answer = Read-MenuValue "$Message | [1] Confirmar | [2] Cancelar"
        if ($null -eq $answer -or $answer -eq '2') { return $false }
        if ($answer -eq '1') { return $true }
        Write-Host 'Resposta invalida. Digite 1 ou 2.' -ForegroundColor Yellow
    }
}

function Register-ReviewSheet([string]$Path, [string]$Type, [int]$Count) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'Planilha de revisao nao encontrada para registrar.' }
    if ($null -eq $script:SessionReviewSheets) { $script:SessionReviewSheets=@() }
    $script:SessionReviewSheets+= [pscustomobject]@{ Path=$Path; Type=$Type; Count=$Count; Created=(Get-Date) }
}

function New-RunDirectory([string]$Name) {
    $path = Join-Path (Join-Path $script:AppRoot 'relatorio') ('{0}-{1}-{2}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'),$Name,[guid]::NewGuid().ToString('N').Substring(0,8))
    $null = New-Item -ItemType Directory -Path $path -ErrorAction Stop
    return $path
}

function Export-Report([object[]]$Rows, [string]$Path) {
    if ($Rows.Count) { $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop }
    else { $Path += '.vazio.txt'; 'Nenhum registro encontrado.' | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop }
    Write-Host "Resultado ($($Rows.Count) registros): $Path" -ForegroundColor Cyan
}

function Import-AdminCsv([string]$Path, [string[]]$Required) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf) -or [IO.Path]::GetExtension($Path) -ine '.csv') { throw 'Informe um CSV existente.' }
    $header = Get-Content -LiteralPath $Path -Encoding UTF8 -TotalCount 1 -ErrorAction Stop
    $delimiter = if ($header -match ';') { ';' } else { ',' }
    $rows = @(Import-Csv -LiteralPath $Path -Delimiter $delimiter -Encoding UTF8 -ErrorAction Stop)
    if (-not $rows.Count) { throw 'CSV sem registros.' }
    foreach ($column in $Required) {
        if ($rows[0].PSObject.Properties.Name -notcontains $column) { throw "Coluna obrigatoria ausente: $column" }
    }
    return $rows
}

function Assert-Smtp([string]$Address) {
    if ($Address -notmatch '^[a-zA-Z0-9._+%\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$') { throw "Endereco SMTP invalido: $Address" }
}

function Get-DisabledAddress([string]$Address) {
    Assert-Smtp $Address
    if ($Address -match '\.desativado@') { return $Address.ToLowerInvariant() }
    return ($Address -replace '@','.desativado@').ToLowerInvariant()
}

function Get-UpdatedAddresses($Mailbox, [string]$NewPrimary) {
    Assert-Smtp $NewPrimary
    $addresses = @("SMTP:$NewPrimary")
    $addresses += @($Mailbox.EmailAddresses | ForEach-Object {
        $address = [string]$_
        if ($address -ine "smtp:$NewPrimary") { $address -creplace '^SMTP:','smtp:' }
    })
    if ([string]$Mailbox.PrimarySmtpAddress -ine $NewPrimary) { $addresses += "smtp:$($Mailbox.PrimarySmtpAddress)" }
    return @($addresses | Select-Object -Unique)
}

function Test-ProxyAddress($Mailbox, [string]$Address) {
    return @($Mailbox.EmailAddresses | ForEach-Object { [string]$_ } |
        Where-Object { $_ -ieq "smtp:$Address" }).Count -gt 0
}

function Assert-AddressAvailable([string]$Address, [string]$ObjectId) {
    Assert-Smtp $Address
    # An empty successful query means free; permissions/network errors must propagate.
    $matches = @(Get-Recipient -Filter "EmailAddresses -eq 'smtp:$Address'" -ResultSize Unlimited -ErrorAction Stop)
    foreach ($match in $matches) {
        if (-not $ObjectId -or [string]$match.ExternalDirectoryObjectId -ine $ObjectId) { throw "Endereco em uso por outro objeto: $Address" }
    }
}

function Ensure-ExchangeConnection {
    Import-Module ExchangeOnlineManagement -Global -ErrorAction Stop
    $connections = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object State -eq 'Connected')
    if (-not $connections.Count) {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $connections = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object State -eq 'Connected')
    }
    $tenants = @($connections | Select-Object -ExpandProperty TenantID -Unique)
    if ($tenants.Count -ne 1 -or -not $tenants[0]) { throw 'Nao foi possivel identificar um unico tenant Exchange. Use uma janela dedicada.' }
    $script:TenantId = [string]$tenants[0]
    $connections | Format-Table UserPrincipalName,Organization,TenantID,State -AutoSize | Out-Host
}

function Ensure-GraphConnection([string[]]$Scopes = @('User.Read.All')) {
    Import-Module Microsoft.Graph.Authentication -Global -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -Global -ErrorAction Stop
    $context = Get-MgContext
    if ($context -and [string]$context.TenantId -ine $script:TenantId) { throw 'Graph e Exchange conectados a tenants diferentes. Corrija a sessao antes de continuar.' }
    $missing = @($Scopes | Where-Object { -not $context -or $context.Scopes -notcontains $_ })
    if (-not $context -or $missing.Count) {
        Connect-MgGraph -TenantId $script:TenantId -Scopes $Scopes -ContextScope Process -NoWelcome -ErrorAction Stop
    }
    $context = Get-MgContext
    if (-not $context -or [string]$context.TenantId -ine $script:TenantId) { throw 'Tenant Graph nao validado.' }
    foreach ($scope in $Scopes) { if ($context.Scopes -notcontains $scope) { throw "Permissao Graph ausente: $scope" } }
}

function Get-AdminInventory {
    $users = @(Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Mail,AccountEnabled,AssignedLicenses -ErrorAction Stop)
    $boxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties ExternalDirectoryObjectId,RecipientTypeDetails,PrimarySmtpAddress,UserPrincipalName,HiddenFromAddressListsEnabled,WhenMailboxCreated -ErrorAction Stop)
    $byId = @{}
    foreach ($box in $boxes) { if ($box.ExternalDirectoryObjectId) { $byId[[string]$box.ExternalDirectoryObjectId] = $box } }
    foreach ($user in $users) {
        [pscustomobject]@{ User=$user; Mailbox=$byId[[string]$user.Id] }
    }
}

function Get-DirectLicenseIds($User) {
    $states = @($User.LicenseAssignmentStates)
    foreach ($license in @($User.AssignedLicenses)) {
        $skuStates = @($states | Where-Object { $_.SkuId -eq $license.SkuId })
        if (-not $skuStates.Count) { throw 'Origem da licenca nao identificada. Remocao bloqueada.' }
    }
    @($states | Where-Object { -not $_.AssignedByGroup -and $_.SkuId } | ForEach-Object { [string]$_.SkuId } | Select-Object -Unique)
}
