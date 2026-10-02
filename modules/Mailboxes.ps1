# SPDX-License-Identifier: MIT
# Copyright (c) 2026 GuilhermeHOliveir
function Assert-LicenseRemovalSafe($Mailbox) {
    # Fail closed if required licensing evidence is unavailable.
    foreach ($property in @('LitigationHoldEnabled','ArchiveStatus','InPlaceHolds','ComplianceTagHoldApplied','DelayHoldApplied','DelayReleaseHoldApplied')) {
        if ($null -eq $Mailbox.$property) { throw "Nao foi possivel verificar $property. Preserve as licencas e revise manualmente." }
    }
    if ($Mailbox.ArchiveStatus -ne 'None' -or (Test-TrueValue $Mailbox.LitigationHoldEnabled) -or
        (Test-TrueValue $Mailbox.ComplianceTagHoldApplied) -or (Test-TrueValue $Mailbox.DelayHoldApplied) -or
        (Test-TrueValue $Mailbox.DelayReleaseHoldApplied) -or @($Mailbox.InPlaceHolds).Count) {
        throw 'Arquivo/hold presente. Remocao automatica de licencas bloqueada.'
    }
    $organization = Get-OrganizationConfig -ErrorAction Stop
    if ($null -eq $organization.InPlaceHolds -or @($organization.InPlaceHolds).Count) { throw 'Holds organizacionais presentes ou nao verificaveis; revise o licenciamento manualmente.' }
    $stats=Get-MailboxStatistics -Identity $Mailbox.Guid -ErrorAction Stop
    $size=$stats.TotalItemSize
    $bytes=$null
    if ($size -and $size.Value -and $size.Value.PSObject.Methods.Name -contains 'ToBytes') { $bytes=[decimal]$size.Value.ToBytes() }
    elseif ([string]$size -match '\(([\d.,\s]+)\s+bytes\)') { $bytes=[decimal]($matches[1] -replace '\D','') }
    if ($null -eq $bytes) { throw 'Tamanho da caixa nao verificavel; remocao de licencas bloqueada.' }
    if ($bytes -ge 50GB) { throw 'Caixa com 50 GB ou mais; preserve o licenciamento.' }
}

function Get-MailboxChangePlan($Row, [bool]$Convert, [bool]$RemoveLicenses) {
    Assert-SpecificIdentity ([string]$Row.UserPrincipalName)
    Assert-Smtp ([string]$Row.PrimaryAtual)
    if ([string]$Row.NovoPrimary -ine (Get-DisabledAddress ([string]$Row.PrimaryAtual))) { throw 'NovoPrimary deve ser o endereco atual com sufixo .desativado.' }
    $user=Get-MgUser -UserId $Row.UserPrincipalName -Property Id,UserPrincipalName,AssignedLicenses,LicenseAssignmentStates -ErrorAction Stop
    $box=Get-Mailbox -Identity $Row.UserPrincipalName -ErrorAction Stop
    if (-not $user.Id -or [string]$box.ExternalDirectoryObjectId -ine [string]$user.Id -or -not $box.Guid) { throw 'Identidade Exchange/Graph divergente ou sem identificador estavel.' }
    if (Test-TrueValue $box.IsDirSynced) { throw 'Objeto sincronizado. Trate a alteracao na origem do ambiente hibrido.' }
    $currentPrimary=[string]$box.PrimarySmtpAddress
    $alreadyRenamed=$currentPrimary -ieq [string]$Row.NovoPrimary
    if ($currentPrimary -ine [string]$Row.PrimaryAtual -and -not $alreadyRenamed) {
        throw 'PrimaryAtual mudou para outro endereco desde o CSV; gere uma nova previa.'
    }
    if ($alreadyRenamed -and -not (Test-ProxyAddress $box ([string]$Row.PrimaryAtual))) {
        throw 'O novo principal ja esta em uso, mas o endereco antigo nao consta como alias. Revise manualmente antes de continuar.'
    }
    if ($Convert) {
        if ($box.RecipientTypeDetails -notin @('UserMailbox','SharedMailbox')) { throw 'Tipo de caixa nao suportado.' }
    } elseif ($box.RecipientTypeDetails -ne 'SharedMailbox' -or @($user.AssignedLicenses | Where-Object { $null -ne $_ }).Count) { throw 'A caixa deve continuar SharedMailbox e sem licencas.' }
    Assert-AddressAvailable ([string]$Row.NovoPrimary) ([string]$user.Id)
    $skus=@()
    if ($RemoveLicenses) {
        $skus=@(Get-DirectLicenseIds $user)
        if ($skus.Count) { Assert-LicenseRemovalSafe $box }
    }
    [pscustomobject]@{ Row=$Row; Guid=[string]$box.Guid; UserId=[string]$user.Id; Type=[string]$box.RecipientTypeDetails
        OldPrimary=$currentPrimary; OriginalPrimary=[string]$Row.PrimaryAtual; NewPrimary=[string]$Row.NovoPrimary; AlreadyRenamed=$alreadyRenamed
        OldAddresses=@($box.EmailAddresses | ForEach-Object { [string]$_ }); Addresses=@(Get-UpdatedAddresses $box ([string]$Row.NovoPrimary))
        DirectSkus=$skus; AssignedLicenses=@($user.AssignedLicenses); LicenseStates=@($user.LicenseAssignmentStates)
        GroupSkus=@($user.LicenseAssignmentStates | Where-Object AssignedByGroup | ForEach-Object { [string]$_.SkuId } | Select-Object -Unique) }
}

function Wait-MailboxState {
    param(
        [string]$Identity,
        [string]$UserId,
        [string]$NewPrimary,
        [string]$OldPrimary,
        [ValidateRange(1,60)][int]$Attempts=13,
        [ValidateRange(1,10)][int]$IntervalSeconds=5
    )
    for ($attempt=1; $attempt -le $Attempts; $attempt++) {
        # Retry only reads, never repeat Set-Mailbox or ignore query failures.
        $box=Get-Mailbox -Identity $Identity -ErrorAction Stop
        if ([string]$box.ExternalDirectoryObjectId -ine $UserId) { throw 'Identidade divergente ao confirmar a caixa.' }
        $valid=$box.RecipientTypeDetails -eq 'SharedMailbox'
        if ($NewPrimary) {
            $valid=$valid -and [string]$box.PrimarySmtpAddress -ieq $NewPrimary -and
                (Test-ProxyAddress $box $OldPrimary)
        }
        if ($valid) { return $box }
        if ($attempt -lt $Attempts) {
            $label=if ($NewPrimary) { 'enderecos da SharedMailbox' } else { 'conversao para SharedMailbox' }
            Write-MenuText "Aguardando confirmacao de $label ($attempt/$Attempts). Nova consulta em $IntervalSeconds segundos..." -ForegroundColor Yellow
            Start-Sleep -Seconds $IntervalSeconds
        }
    }
    $detail="Tipo consultado: $($box.RecipientTypeDetails); principal: $($box.PrimarySmtpAddress)."
    throw [TimeoutException]::new("Alteracao enviada, mas ainda nao confirmada apos $Attempts consultas. $detail Licencas preservadas; consulte o estado antes de repetir.")
}

function Wait-DirectLicensesRemoved {
    param(
        [string]$UserId,
        [string[]]$SkuIds,
        [ValidateRange(1,60)][int]$Attempts=7,
        [ValidateRange(1,10)][int]$IntervalSeconds=5
    )
    if (-not $UserId -or -not $SkuIds.Count) { throw 'Usuario ou licencas ausentes na verificacao.' }
    for ($attempt=1; $attempt -le $Attempts; $attempt++) {
        $user=Get-MgUser -UserId $UserId -Property Id,AssignedLicenses,LicenseAssignmentStates -ErrorAction Stop
        if ([string]$user.Id -ine $UserId) { throw 'Identidade Graph divergente ao verificar licencas.' }
        $consistent=$true
        try { $direct=@(Get-DirectLicenseIds $user) }
        catch {
            if ($_.Exception.Message -ne 'Origem da licenca nao identificada. Remocao bloqueada.') { throw }
            # Graph may briefly return assignedLicenses and assignmentStates from different revisions.
            $consistent=$false
        }
        if ($consistent -and -not @($direct | Where-Object { $SkuIds -contains $_ }).Count) { return $user }
        if ($attempt -lt $Attempts) {
            Write-MenuText "Aguardando confirmacao de licencas ($attempt/$Attempts). Nova consulta em $IntervalSeconds segundos..." -ForegroundColor Yellow
            Start-Sleep -Seconds $IntervalSeconds
        }
    }
    throw [TimeoutException]::new("Remocao enviada, mas nao confirmada apos $Attempts consultas. Consulte as licencas antes de repetir.")
}

function Invoke-MailboxChanges([string]$Path, [bool]$Convert, [bool]$RemoveLicenses, [bool]$Apply = $false, [switch]$PassThru) {
    $ErrorActionPreference='Stop'
    $rows=@(Import-AdminCsv $Path @('Processar','UserPrincipalName','PrimaryAtual','NovoPrimary') | Where-Object Processar -eq 'SIM')
    if (-not $rows.Count) { throw 'Nenhum destino marcado Processar=SIM.' }
    $plans=@(); $seen=@{}
    # Validate the whole CSV before the first remote write.
    foreach ($row in $rows) {
        $plan=Get-MailboxChangePlan $row $Convert $RemoveLicenses
        if ($seen.ContainsKey($plan.Guid)) { throw 'CSV possui destinos duplicados. Remova as duplicatas antes de continuar.' }
        $seen[$plan.Guid]=$true; $plans+= $plan
    }
    Invoke-MailboxPlan -Plans $plans -Convert $Convert -RemoveLicenses $RemoveLicenses -Apply $Apply -PassThru:$PassThru
}

function Invoke-SharedRenameInteractive {
    $inventory=@(Get-AdminInventory)
    $plans=@(); $seen=@{}; $alreadyRenamed=0; $blocked=0
    foreach ($entry in $inventory) {
        $user=$entry.User; $box=$entry.Mailbox
        if (-not $box -or $box.RecipientTypeDetails -ne 'SharedMailbox' -or @($user.AssignedLicenses | Where-Object { $null -ne $_ }).Count) { continue }
        $old=[string]$box.PrimarySmtpAddress
        if ($old -match '\.desativado@') {
            $alreadyRenamed++
            continue
        }
        try {
            $new=Get-DisabledAddress $old
            $row=[pscustomobject]@{ UserPrincipalName=[string]$user.UserPrincipalName; PrimaryAtual=$old; NovoPrimary=$new }
            $plan=Get-MailboxChangePlan $row $false $false
            if ($seen.ContainsKey($plan.Guid)) { throw 'Conta duplicada no inventario.' }
            $seen[$plan.Guid]=$true
            $plans+=$plan
        } catch {
            $blocked++
        }
    }
    Show-Header 'SHAREDMAILBOXES QUE SERAO ALTERADAS'
    for ($i=0; $i -lt $plans.Count; $i++) {
        $plan=$plans[$i]
        Write-MenuText ("[{0}] {1}" -f ($i+1),$plan.Row.UserPrincipalName) -ForegroundColor Cyan
        Write-MenuText "    Atual: $($plan.OriginalPrimary)"
        Write-MenuText "    Novo:  $($plan.NewPrimary)"
    }
    Write-MenuText "Elegiveis: $($plans.Count) | Ja com .desativado: $alreadyRenamed | Impedidas: $blocked" -ForegroundColor Cyan
    if (-not $plans.Count) { Write-MenuText 'Nenhuma caixa para alterar.' -ForegroundColor Yellow; return }
    Write-MenuText 'O endereco antigo permanecera como alias. O sufixo .desativado nao bloqueia login nem recebimento.' -ForegroundColor Yellow
    while ($true) {
        $answer=Read-MenuValue "Alterar EM PRODUCAO as $($plans.Count) caixa(s) prontas acima? [1] Aplicar | [0] Cancelar"
        if ($null -eq $answer) { Write-MenuText 'Alteracao cancelada.'; return }
        if ($answer -eq '1') { break }
        Write-MenuText 'Digite 1 para aplicar ou 0 para cancelar.' -ForegroundColor Yellow
    }
    Invoke-MailboxPlan -Plans $plans -Convert $false -RemoveLicenses $false -Apply $true -Confirmed
    Wait-Menu
}

function Invoke-MailboxPlan([object[]]$Plans, [bool]$Convert, [bool]$RemoveLicenses, [bool]$Apply=$false, [switch]$Confirmed, [switch]$PassThru) {
    $ErrorActionPreference='Stop'
    if (-not $Plans.Count) { throw 'Plano vazio. Operacao bloqueada.' }
    $plans | Select-Object OldPrimary,NewPrimary,Type,@{n='LicencasDiretas';e={$_.DirectSkus.Count}},@{n='LicencasGrupo';e={$_.GroupSkus.Count}} | Out-MenuTable -Wrap -AutoSize
    if ($RemoveLicenses) {
        Write-MenuText 'Remove TODOS os produtos/SKUs diretos selecionados, podendo afetar OneDrive, Teams e outros servicos. Licencas de grupo permanecem. Revise tambem recursos licenciados do Purview/Defender.' -ForegroundColor Yellow
    }
    Write-MenuText 'O sufixo .desativado nao bloqueia login nem entrega de mensagens. O alias antigo permanece.' -ForegroundColor Yellow
    if ($Apply -and -not $Confirmed -and -not (Confirm-Operation "APLICAR em $($plans.Count) caixa(s) acima")) { return }
    $folder=New-RunDirectory 'alteracao-caixas'
    $results=@(foreach ($plan in $plans) {
        $stage='Revalidacao'; $status='Simulado'; $errorText=''
        try {
            $fresh=Get-MailboxChangePlan $plan.Row $Convert $RemoveLicenses
            if ($fresh.Guid -ne $plan.Guid -or $fresh.UserId -ne $plan.UserId -or $fresh.Type -ne $plan.Type -or
                (($fresh.OldAddresses | Sort-Object) -join '|') -cne (($plan.OldAddresses | Sort-Object) -join '|') -or
                (($fresh.DirectSkus | Sort-Object) -join '|') -ne (($plan.DirectSkus | Sort-Object) -join '|')) { throw 'Estado mudou desde a previa. Operacao bloqueada para esta conta.' }
            Write-Audit 'BackupCaixa' $fresh
            if ($Apply) {
                if ($Convert -and $fresh.Type -eq 'UserMailbox') {
                    $stage='Conversao enviada'
                    Set-Mailbox -Identity $fresh.Guid -Type Shared -ErrorAction Stop
                }
                $check=Wait-MailboxState -Identity $fresh.Guid -UserId $fresh.UserId
                if (-not $fresh.AlreadyRenamed) {
                    $stage='Alteracao de enderecos enviada'
                    Set-Mailbox -Identity $fresh.Guid -EmailAddresses $fresh.Addresses -ErrorAction Stop
                }
                $check=Wait-MailboxState -Identity $fresh.Guid -UserId $fresh.UserId -NewPrimary $fresh.NewPrimary -OldPrimary $fresh.OriginalPrimary
                if ($fresh.DirectSkus.Count) {
                    Assert-LicenseRemovalSafe $check
                    $currentUser=Get-MgUser -UserId $fresh.UserId -Property Id,AssignedLicenses,LicenseAssignmentStates -ErrorAction Stop
                    $currentSkus=@(Get-DirectLicenseIds $currentUser)
                    if ((($currentSkus | Sort-Object) -join '|') -ne (($fresh.DirectSkus | Sort-Object) -join '|')) { throw 'Licencas mudaram durante a execucao. Remocao bloqueada.' }
                    $stage='Remocao de licencas enviada'
                    $null=Set-MgUserLicense -UserId $fresh.UserId -AddLicenses @() -RemoveLicenses $fresh.DirectSkus -ErrorAction Stop
                    $null=Wait-DirectLicensesRemoved -UserId $fresh.UserId -SkuIds $fresh.DirectSkus
                }
                $stage='Validado'; $status='Alterado e validado'
            }
        } catch [TimeoutException] {
            $status='Pendente de confirmacao'; $errorText=$_.Exception.Message
        } catch { $status='Erro / revisar estado'; $errorText=$_.Exception.Message }
        $result=[pscustomobject]@{ UserPrincipalName=$plan.Row.UserPrincipalName; Guid=$plan.Guid; PrimaryAnterior=$plan.OldPrimary
            NovoPrimary=$plan.NewPrimary; Etapa=$stage; Status=$status; Erro=$errorText }
        $result | Export-Csv -LiteralPath (Join-Path $folder 'resultado.csv') -NoTypeInformation -Encoding UTF8 -Append -ErrorAction Stop
        Write-Audit 'ResultadoCaixa' $result
        $result
    })
    foreach ($result in $results) {
        Write-MenuText "`nConta: $($result.UserPrincipalName)" -ForegroundColor Cyan
        Write-MenuText "  Etapa:  $($result.Etapa)"
        Write-MenuText "  Estado: $($result.Status)"
        if ($result.Erro) { Write-MenuText "  Detalhe: $($result.Erro)" -ForegroundColor Yellow }
    }
    $results | Group-Object Status | Select-Object Name,Count | Out-MenuTable -AutoSize
    Write-MenuText "Resultados: $folder | Backup anterior: $script:LogPath"
    if ($PassThru) {
        [pscustomobject]@{ Plans=$Plans; Results=$results; Convert=$Convert; RemoveLicenses=$RemoveLicenses; Applied=$Apply }
    }
}

function Complete-MailboxSimulation($Simulation) {
    if (-not $Simulation -or $Simulation.Applied -or -not @($Simulation.Plans).Count -or
        @($Simulation.Results).Count -ne @($Simulation.Plans).Count -or
        @($Simulation.Results | Where-Object Status -ne 'Simulado').Count) {
        Write-MenuText 'Simulacao com falhas ou incompleta. Corrija os erros antes de aplicar.' -ForegroundColor Yellow
        Wait-Menu
        return
    }
    if (-not (Confirm-Operation "Simulacao concluida. Deseja aplicar EM PRODUCAO nas mesmas $(@($Simulation.Plans).Count) caixa(s), com as mesmas opcoes")) {
        Write-MenuText 'Processo finalizado sem aplicar em producao.'
        return
    }
    # Keep the simulated plan in memory: never reread a CSV that may have changed.
    Ensure-ExchangeConnection
    $scopes=@('User.Read.All')
    if ($Simulation.RemoveLicenses) {
        $scopes+='LicenseAssignment.ReadWrite.All'
        Ensure-AdminModule Microsoft.Graph.Users.Actions
    }
    Ensure-GraphConnection $scopes
    Invoke-MailboxPlan -Plans $Simulation.Plans -Convert $Simulation.Convert -RemoveLicenses $Simulation.RemoveLicenses -Apply $true -Confirmed
    Wait-Menu
}
