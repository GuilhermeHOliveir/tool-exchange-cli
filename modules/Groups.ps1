# SPDX-License-Identifier: MIT
# Copyright (c) 2026 GuilhermeHOliveir
function Get-EnderecosDoArquivo {
    param(
        [Parameter(Mandatory)][string]$Caminho,
        [Parameter(Mandatory)][string]$NomeColuna
    )

    $extensao = [IO.Path]::GetExtension($Caminho).ToLowerInvariant()

    if ($extensao -eq '.csv') {
        $primeiraLinha = Get-Content -LiteralPath $Caminho |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 1

        if (-not $primeiraLinha) {
            throw "O arquivo CSV esta vazio: $Caminho"
        }

        $delimitador = if (($primeiraLinha.ToCharArray() | Where-Object { $_ -eq ';' }).Count -gt
            ($primeiraLinha.ToCharArray() | Where-Object { $_ -eq ',' }).Count) { ';' } else { ',' }

        $linhas = @(Import-Csv -LiteralPath $Caminho -Delimiter $delimitador)
        if ($linhas.Count -eq 0) {
            throw "O arquivo CSV nao contem registros: $Caminho"
        }

        if ($linhas[0].PSObject.Properties.Name -notcontains $NomeColuna) {
            $colunas = $linhas[0].PSObject.Properties.Name -join ', '
            throw "A coluna '$NomeColuna' nao foi encontrada. Colunas disponiveis: $colunas"
        }

        return @($linhas | ForEach-Object { [string]$_.$NomeColuna })
    }

    if ($extensao -eq '.txt') {
        return @(Get-Content -LiteralPath $Caminho)
    }

    throw "Formato nao aceito: '$extensao'. Use um arquivo .csv ou .txt."
}

function Get-SmtpNormalizado {
    param([AllowNull()][object]$Valor)

    if ($null -eq $Valor) { return $null }
    $texto = ([string]$Valor).Trim()
    if ([string]::IsNullOrWhiteSpace($texto)) { return $null }
    return ($texto -replace '^(?i)smtp:', '').Trim().ToLowerInvariant()
}

function Test-SameRecipient($Member, $Candidate) {
    if ($Member.ExternalDirectoryObjectId -and $Candidate.ExternalDirectoryObjectId) {
        return [string]$Member.ExternalDirectoryObjectId -ieq [string]$Candidate.ExternalDirectoryObjectId
    }
    if ($Member.Guid -and $Candidate.Guid) { return [string]$Member.Guid -eq [string]$Candidate.Guid }
    return ($Member.PrimarySmtpAddress -and $Candidate.PrimarySmtpAddress -and
        [string]$Member.PrimarySmtpAddress -ieq [string]$Candidate.PrimarySmtpAddress)
}

function Invoke-MemberImport([string]$Group, [string]$Path, [string]$Column='Email', [bool]$Apply=$false, [bool]$Bypass=$false) {
    $ErrorActionPreference='Stop'
    Assert-SpecificIdentity $Group
    $groupObject=Get-DistributionGroup -Identity $Group -ErrorAction Stop
    if (-not $groupObject.Guid) { throw 'Grupo sem GUID. Operacao bloqueada.' }
    if (Test-TrueValue $groupObject.IsDirSynced) { throw 'Grupo sincronizado; altere na origem.' }
    $members=@(Get-DistributionGroupMember -Identity $groupObject.Guid -ResultSize Unlimited -ErrorAction Stop)
    $seen=@{}; $plans=@()
    foreach ($raw in @(Get-EnderecosDoArquivo $Path $Column)) {
        $email=Get-SmtpNormalizado $raw
        if (-not $email) { continue }
        Assert-Smtp $email
        $recipient=Get-Recipient -Identity $email -ErrorAction Stop
        if (-not $recipient.Guid) { throw "Destinatario sem GUID: $email" }
        $id=[string]$recipient.Guid
        if ($seen.ContainsKey($id)) { continue }
        $seen[$id]=$true
        $exists=@($members | Where-Object { Test-SameRecipient $_ $recipient }).Count -gt 0
        $plans+=[pscustomobject]@{ Email=$email; Recipient=$recipient; Exists=$exists }
    }
    if (-not $plans.Count) { throw 'Arquivo sem destinatarios validos.' }
    Write-Host "Grupo: $($groupObject.DisplayName) <$($groupObject.PrimarySmtpAddress)>"
    $plans | Select-Object Email,@{n='JaMembro';e={$_.Exists}} | Format-Table -AutoSize | Out-Host
    if ($Apply -and -not (Confirm-Operation "Adicionar membros ausentes ao grupo $($groupObject.PrimarySmtpAddress)")) { return }
    $folder=New-RunDirectory 'importacao-membros'
    foreach ($plan in $plans) {
        $status='Simulado'; $errorText=''
        try {
            if ($plan.Exists) { $status='Ja membro' }
            elseif ($Apply) {
                Write-Audit 'AdicionarMembroSolicitado' @{Group=[string]$groupObject.Guid; Member=[string]$plan.Recipient.Guid; Email=$plan.Email}
                $parameters=@{Identity=$groupObject.Guid; Member=$plan.Recipient.Guid; ErrorAction='Stop'}
                if ($Bypass) { $parameters.BypassSecurityGroupManagerCheck=$true }
                Add-DistributionGroupMember @parameters
                $after=@(Get-DistributionGroupMember -Identity $groupObject.Guid -ResultSize Unlimited -ErrorAction Stop)
                if (-not @($after | Where-Object { Test-SameRecipient $_ $plan.Recipient }).Count) { throw 'Inclusao enviada, mas ainda nao confirmada.' }
                $status='Adicionado e validado'
            }
        } catch { $status='Erro / revisar estado'; $errorText=$_.Exception.Message }
        $result=[pscustomobject]@{Grupo=$groupObject.PrimarySmtpAddress; Email=$plan.Email; Status=$status; Erro=$errorText}
        $result | Export-Csv -LiteralPath (Join-Path $folder 'resultado.csv') -NoTypeInformation -Encoding UTF8 -Append -ErrorAction Stop
        Write-Audit 'ResultadoImportacao' $result
        $result | Format-Table -AutoSize | Out-Host
    }
    Write-Host "Resultados: $folder"
}

function Get-DisabledGroupCandidates {
    foreach ($entry in @(Get-AdminInventory)) {
        $box=$entry.Mailbox
        if ($box -and $box.RecipientTypeDetails -eq 'SharedMailbox' -and @($entry.User.AssignedLicenses).Count -eq 0 -and [string]$box.PrimarySmtpAddress -match '\.desativado@') {
            [pscustomobject]@{ ExternalDirectoryObjectId=[string]$entry.User.Id; UserPrincipalName=$entry.User.UserPrincipalName
                PrimarySmtpAddress=[string]$box.PrimarySmtpAddress; DisplayName=$box.DisplayName }
        }
    }
}

function Assert-StillDisabled($Candidate) {
    $user=Get-MgUser -UserId $Candidate.ExternalDirectoryObjectId -Property Id,AssignedLicenses -ErrorAction Stop
    $box=Get-Mailbox -Identity $Candidate.PrimarySmtpAddress -ErrorAction Stop
    if ([string]$box.ExternalDirectoryObjectId -ine [string]$user.Id -or $box.RecipientTypeDetails -ne 'SharedMailbox' -or
        @($user.AssignedLicenses).Count -ne 0 -or [string]$box.PrimarySmtpAddress -notmatch '\.desativado@') { throw 'Conta deixou de atender aos criterios da previa.' }
    return $box
}

function Get-GroupCleanupPlan([object[]]$Candidates, [bool]$RemoveOwners=$false) {
    foreach ($kind in @('Distribution','Unified')) {
        $groups=if ($kind -eq 'Distribution') { @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop) } else { @(Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop) }
        foreach ($group in $groups) {
            if (-not $group.Guid) { throw 'Grupo sem GUID. Consulta interrompida.' }
            if ($kind -eq 'Distribution') {
                $members=@(Get-DistributionGroupMember -Identity $group.Guid -ResultSize Unlimited -ErrorAction Stop); $owners=@()
            } else {
                $members=@(Get-UnifiedGroupLinks -Identity $group.Guid -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
                $owners=@(Get-UnifiedGroupLinks -Identity $group.Guid -LinkType Owners -ResultSize Unlimited -ErrorAction Stop)
            }
            foreach ($candidate in $Candidates) {
                $isOwner=@($owners | Where-Object { Test-SameRecipient $_ $candidate }).Count -gt 0
                foreach ($role in @('Owners','Members')) {
                    $collection=if ($role -eq 'Owners') { $owners } else { $members }
                    if (-not @($collection | Where-Object { Test-SameRecipient $_ $candidate }).Count) { continue }
                    $blocked=''
                    if (Test-TrueValue $group.IsDirSynced) { $blocked='Grupo sincronizado' }
                    elseif ($isOwner -and -not $RemoveOwners) { $blocked='Owner preservado (inclusive vinculo de membro)' }
                    elseif ($isOwner -and $owners.Count -le 1) { $blocked='Ultimo owner preservado' }
                    $groupId=[string]$group.Guid
                    $address=([string]$candidate.PrimarySmtpAddress).Replace("'","''")
                    $restore=if ($kind -eq 'Distribution') { "Add-DistributionGroupMember -Identity '$groupId' -Member '$address'" } else { "Add-UnifiedGroupLinks -Identity '$groupId' -LinkType $role -Links '$address'" }
                    [pscustomobject]@{ Candidate=$candidate; GroupId=$groupId; GroupName=$group.DisplayName; Kind=$kind; Role=$role; Blocked=$blocked; RestoreCommand=$restore }
                }
            }
        }
    }
}

function Invoke-GroupCleanup([bool]$Apply=$false, [bool]$RemoveOwners=$false) {
    $ErrorActionPreference='Stop'
    $candidates=@(Get-DisabledGroupCandidates)
    if (-not $candidates.Count) { Write-Host 'Nenhuma SharedMailbox sem licenca com .desativado encontrada.'; return }
    $plans=@(Get-GroupCleanupPlan $candidates $RemoveOwners)
    if (-not $plans.Count) { Write-Host 'Nenhum vinculo encontrado.'; return }
    $plans | Select-Object @{n='Conta';e={$_.Candidate.PrimarySmtpAddress}},GroupName,Role,Blocked | Format-Table -Wrap -AutoSize | Out-Host
    Write-Host 'Escopo: todas as contas acima em distribuicao/seguranca habilitada para email e Microsoft 365 (inclusive Teams). Grupos dinamicos podem recusar alteracao manual.' -ForegroundColor Yellow
    if ($Apply -and -not (Confirm-Operation 'Remover os vinculos elegiveis acima')) { return }
    $folder=New-RunDirectory 'remocao-grupos'
    Export-Report $candidates (Join-Path $folder 'candidatos.csv')
    foreach ($plan in $plans) {
        $status='Simulado'; $errorText=''
        $backup=[pscustomobject]@{DataHora=(Get-Date).ToString('o'); UserPrincipalName=$plan.Candidate.UserPrincipalName
            PrimarySmtpAddress=$plan.Candidate.PrimarySmtpAddress; GraphUserId=$plan.Candidate.ExternalDirectoryObjectId
            GroupIdentity=$plan.GroupId; GroupName=$plan.GroupName; GroupType=$plan.Kind; Role=$plan.Role; RestoreCommand=$plan.RestoreCommand}
        $backup | Export-Csv -LiteralPath (Join-Path $folder 'backup-vinculos.csv') -NoTypeInformation -Encoding UTF8 -Append -ErrorAction Stop
        try {
            if ($plan.Blocked) { $status='Preservado'; $errorText=$plan.Blocked }
            else {
                $box=Assert-StillDisabled $plan.Candidate
                if ($Apply) {
                    Write-Audit 'RemoverVinculoSolicitado' $backup
                    if ($plan.Kind -eq 'Distribution') {
                        $groupNow=Get-DistributionGroup -Identity $plan.GroupId -ErrorAction Stop
                        if (Test-TrueValue $groupNow.IsDirSynced) { throw 'Grupo tornou-se sincronizado.' }
                        Remove-DistributionGroupMember -Identity $plan.GroupId -Member $box.Guid -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                        $remaining=@(Get-DistributionGroupMember -Identity $plan.GroupId -ResultSize Unlimited -ErrorAction Stop)
                    } else {
                        $owners=@(Get-UnifiedGroupLinks -Identity $plan.GroupId -LinkType Owners -ResultSize Unlimited -ErrorAction Stop)
                        $isOwner=@($owners | Where-Object { Test-SameRecipient $_ $plan.Candidate }).Count -gt 0
                        if ($isOwner -and ($plan.Role -eq 'Members' -or -not $RemoveOwners -or $owners.Count -le 1)) { throw 'Owner preservado: nao autorizado, ultimo owner ou ownership ainda nao removido.' }
                        Remove-UnifiedGroupLinks -Identity $plan.GroupId -LinkType $plan.Role -Links $box.Guid -Confirm:$false -ErrorAction Stop
                        $remaining=@(Get-UnifiedGroupLinks -Identity $plan.GroupId -LinkType $plan.Role -ResultSize Unlimited -ErrorAction Stop)
                    }
                    if (@($remaining | Where-Object { Test-SameRecipient $_ $plan.Candidate }).Count) { throw 'Remocao enviada, mas ainda nao confirmada.' }
                    $status='Removido e validado'
                }
            }
        } catch { $status='Erro / revisar estado'; $errorText=$_.Exception.Message }
        $result=[pscustomobject]@{Conta=$plan.Candidate.PrimarySmtpAddress; Grupo=$plan.GroupName; Role=$plan.Role; Status=$status; Erro=$errorText}
        $result | Export-Csv -LiteralPath (Join-Path $folder 'resultado.csv') -NoTypeInformation -Encoding UTF8 -Append -ErrorAction Stop
        Write-Audit 'ResultadoRemocaoVinculo' $result
        $result | Format-Table -Wrap -AutoSize | Out-Host
    }
    Write-Host "Resultados e comandos de restauracao: $folder"
}
