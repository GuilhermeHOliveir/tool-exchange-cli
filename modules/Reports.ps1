function Invoke-UnlicensedReport {
    $folder = New-RunDirectory 'sem-licenca'
    $rows = @(foreach ($entry in @(Get-AdminInventory)) {
        $user = $entry.User; $box = $entry.Mailbox
        if (@($user.AssignedLicenses).Count -eq 0 -and $box -and $box.RecipientTypeDetails -ne 'SharedMailbox') {
            [pscustomobject]@{ DisplayName=$user.DisplayName; UserPrincipalName=$user.UserPrincipalName; Mail=$user.Mail
                AccountEnabled=$user.AccountEnabled; Licenciado=$false; TipoMailbox=$box.RecipientTypeDetails
                PrimarySmtpAddress=$box.PrimarySmtpAddress; HiddenFromAddressListsEnabled=$box.HiddenFromAddressListsEnabled
                WhenMailboxCreated=$box.WhenMailboxCreated; Observacao='Usuario sem licenca possui mailbox diferente de SharedMailbox' }
        }
    })
    Export-Report $rows (Join-Path $folder 'usuarios_sem_licenca_nao_shared.csv')
    $rows | Format-Table DisplayName,PrimarySmtpAddress,TipoMailbox -AutoSize | Out-Host
}

function Invoke-SharedPreview {
    $folder = New-RunDirectory 'shared-pendente'
    $rows = @(foreach ($entry in @(Get-AdminInventory)) {
        $user=$entry.User; $box=$entry.Mailbox
        if (@($user.AssignedLicenses | Where-Object { $null -ne $_ }).Count -ne 0 -or -not $box -or $box.RecipientTypeDetails -ne 'SharedMailbox') { continue }
        $old = [string]$box.PrimarySmtpAddress
        if ($old -match '\.desativado@') { continue }
        $new = Get-DisabledAddress $old
        $reason = 'OK para alterar'; $conflict = $false
        try { Assert-AddressAvailable $new ([string]$user.Id) } catch { $conflict=$true; $reason=$_.Exception.Message }
        [pscustomobject]@{ Processar='NAO'; DisplayName=$box.DisplayName; UserPrincipalName=$user.UserPrincipalName
            AccountEnabled=$user.AccountEnabled; TipoMailbox='SharedMailbox'; PrimaryAtual=$old; NovoPrimary=$new
            ManterAliasOriginal=$old; NovoPrimaryJaExiste=$conflict; HiddenFromAddressListsEnabled=$box.HiddenFromAddressListsEnabled; Observacao=$reason }
    })
    $path=Join-Path $folder 'preview_shared_desativado_pendente.csv'
    Export-Report $rows $path
    if ($rows.Count) { Register-ReviewSheet $path 'SharedMailbox sem licenca' $rows.Count }
    Write-Host 'Revise o CSV e marque Processar=SIM somente nos destinos aprovados.' -ForegroundColor Yellow
}

function Invoke-InactivityReport {
    Import-Module Microsoft.Graph.Reports -Global -ErrorAction Stop
    $usage=@{}; $activity=@{}
    $tempFolder=Join-Path ([IO.Path]::GetTempPath()) ('exchange-admin-'+[guid]::NewGuid().ToString('N'))
    $null=New-Item -ItemType Directory -Path $tempFolder -ErrorAction Stop
    $usagePath=Join-Path $tempFolder 'mailbox_usage_D90.csv'
    $activityPath=Join-Path $tempFolder 'email_activity_D90.csv'
    try {
        Get-MgReportMailboxUsageDetail -Period D90 -OutFile $usagePath -ErrorAction Stop
        Get-MgReportEmailActivityUserDetail -Period D90 -OutFile $activityPath -ErrorAction Stop
        foreach ($row in @(Import-Csv -LiteralPath $usagePath -ErrorAction Stop)) { if ($row.'User Principal Name') { $usage[$row.'User Principal Name']=$row } }
        foreach ($row in @(Import-Csv -LiteralPath $activityPath -ErrorAction Stop)) { if ($row.'User Principal Name') { $activity[$row.'User Principal Name']=$row } }
    } finally {
        foreach ($file in @($usagePath,$activityPath)) { if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue } }
        if (Test-Path -LiteralPath $tempFolder) { Remove-Item -LiteralPath $tempFolder -ErrorAction SilentlyContinue }
    }
    $exceptions=@()
    $configPath=Join-Path $script:AppRoot 'config.local.json'
    if (Test-Path -LiteralPath $configPath) { $exceptions=@((Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop).ExcecoesSuspeitos) }
    $rows=@(foreach ($entry in @(Get-AdminInventory)) {
        $user=$entry.User; $box=$entry.Mailbox
        if (@($user.AssignedLicenses).Count -eq 0) { continue }
        $upn=[string]$user.UserPrincipalName; $a=$activity[$upn]; $u=$usage[$upn]
        $last=$null; $days=$null; $send=$null; $receive=$null; $read=$null
        $countsValid=$null -ne $a
        foreach ($field in @('Send Count','Receive Count','Read Count')) {
            $number=0
            if (-not [int]::TryParse([string]$a.$field,[ref]$number) -or $number -lt 0) { $countsValid=$false }
        }
        if ($countsValid) { $send=[int]$a.'Send Count'; $receive=[int]$a.'Receive Count'; $read=[int]$a.'Read Count' }
        $dates=@(foreach ($value in @($a.'Last Activity Date',$u.'Last Activity Date')) {
            if ($value) { [datetime]::Parse($value,[Globalization.CultureInfo]::InvariantCulture) }
        })
        if ($dates.Count) { $last=$dates | Sort-Object -Descending | Select-Object -First 1; $days=((Get-Date)-$last).Days }
        $status='Ativo'; $reason='Possui atividade recente de email'
        if (-not $box -or $box.RecipientTypeDetails -ne 'UserMailbox') { $status='Ignorado'; $reason='Sem UserMailbox correspondente' }
        elseif (-not $a -or -not $u -or -not $countsValid) { $status='Dados insuficientes'; $reason='Relatorio ausente/anonimizado/incompleto; ausencia de dados nao comprova inatividade' }
        elseif ($user.AccountEnabled -eq $false) { $status='Validar manualmente'; $reason='Conta desabilitada, mas licenciada' }
        elseif (($null -eq $days -or $days -ge 90) -and $send -eq 0 -and $read -eq 0) {
            $status='Validar manualmente'; $reason='Sem envio/leitura em D90; login nao consultado'
        }
        [pscustomobject]@{ DisplayName=$user.DisplayName; UserPrincipalName=$upn; Mail=$user.Mail; AccountEnabled=$user.AccountEnabled
            TipoMailbox=$box.RecipientTypeDetails; PrimarySmtpAddress=$box.PrimarySmtpAddress; Licenciado=$true
            QuantidadeLicencas=@($user.AssignedLicenses).Count; LastMailboxActivityDate=$last; SendCount90Dias=$send
            ReceiveCount90Dias=$receive; ReadCount90Dias=$read; LastSignIn=$null; DiasSemLogin=$null
            DiasSemAtividadeEmail=$days; StatusValidacao=$status; Motivo=$reason }
    })
    $suspects=@(foreach ($row in $rows) {
        if ($row.StatusValidacao -ne 'Validar manualmente' -or $exceptions -contains $row.UserPrincipalName -or $exceptions -contains $row.Mail -or $exceptions -contains [string]$row.PrimarySmtpAddress) { continue }
        $result=[ordered]@{ Processar='NAO'; PrimaryAtual=[string]$row.PrimarySmtpAddress; NovoPrimary=(Get-DisabledAddress ([string]$row.PrimarySmtpAddress)); ObservacaoValidacao='' }
        foreach ($p in $row.PSObject.Properties) { $result[$p.Name]=$p.Value }
        [pscustomobject]$result
    })
    $rows | Group-Object StatusValidacao | Select-Object Name,Count | Format-Table -AutoSize | Out-Host
    if (-not $suspects.Count) { Write-Host 'Nenhuma conta elegivel para revisao. Nenhum CSV foi gerado.' -ForegroundColor Yellow; return }
    $folder=New-RunDirectory 'inatividade'
    $path=Join-Path $folder 'emails_suspeitos_para_validar.csv'
    Export-Report $suspects $path
    Register-ReviewSheet $path 'Inatividade D90' $suspects.Count
    Write-Host 'Revise o CSV e marque Processar=SIM somente nas contas aprovadas. Salve antes de iniciar a conversao.' -ForegroundColor Yellow
    Write-Host 'Outros servicos Microsoft 365 e login nao foram avaliados.' -ForegroundColor Yellow
    try { Invoke-Item -LiteralPath $path -ErrorAction Stop }
    catch { Write-Host "Nao foi possivel abrir automaticamente: $($_.Exception.Message)" -ForegroundColor Yellow }
}
