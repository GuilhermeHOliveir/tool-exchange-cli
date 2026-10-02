#requires -Version 5.1
<# Menu de administracao MRM. Ao importar com dot-source, apenas carrega funcoes. #>
[CmdletBinding()]
param(
    [string]$ContasTxt,
    [string]$LogDirectory = (Join-Path $PSScriptRoot 'logs')
)

$script:ArquivoContas = $ContasTxt
$script:LogPath = $null
$script:Conectou = $false

function Test-TrueValue($Value) { return ([string]$Value -eq 'True') }

function Get-RetentionDays($Value) {
    if ($null -ne $Value -and $null -ne $Value.TotalDays) { return [double]$Value.TotalDays }
    $span = [timespan]::Zero
    if ([timespan]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [ref]$span)) {
        return $span.TotalDays
    }
    throw 'O Exchange retornou um prazo de retencao que nao foi possivel interpretar.'
}

function Test-ProtectedPolicy($Policy) {
    return (([string]$Policy.Name -ieq 'Default MRM Policy') -or
        (Test-TrueValue $Policy.IsDefault) -or (Test-TrueValue $Policy.IsDefaultArbitrationMailbox))
}

function Assert-SpecificIdentity([string]$Identity) {
    # Aceita alias ou SMTP/UPN, nunca curingas, filtros ou listas em uma linha.
    if ([string]::IsNullOrWhiteSpace($Identity) -or
        $Identity -notmatch '^[\p{L}\p{N}][\p{L}\p{N}._+@%\-]*$' -or
        $Identity -in @('all', 'todos', 'todas', 'unlimited')) {
        throw "Identidade invalida: '$Identity'. Use um alias ou e-mail exato, sem curingas."
    }
}

function Read-MenuValue([string]$Prompt) {
    if (-not $script:MenuWidth) { Initialize-MenuLayout }
    Write-Host ''
    Write-MenuText ('-' * $script:MenuWidth) DarkCyan
    $navigation = '0  Voltar    |    00  Menu principal    |    000  Sair'
    if ($script:MenuTitle -eq 'MENU PRINCIPAL') { $navigation = '0 / 000  Encerrar    |    Digite o numero da opcao desejada' }
    Write-MenuText $navigation DarkGray
    Write-Host ''
    Write-MenuText ($Prompt -replace ' \| ', "`n") White
    $value = (Read-Host ($script:MenuIndent + '  >')).Trim()
    if ($value -ceq '000') { throw 'MRM_MENU_EXIT' }
    if ($value -ceq '00') { throw 'MRM_MENU_HOME' }
    if ($value -eq '0') { return $null }
    return $value
}

function Clear-MenuScreen {
    # Use the native console clear instead of relying on the host's Clear-Host
    # implementation. Clear the buffer before repositioning and drawing a menu.
    if ($Host.Name -eq 'ConsoleHost' -and -not [Console]::IsOutputRedirected) {
        try {
            [Console]::Clear()
            [Console]::SetCursorPosition(0, 0)
            return
        } catch {
            # ISE, redirected sessions and hosts without a console use their API.
        }
    }
    Clear-Host
}

function Initialize-MenuLayout {
    $width = 80
    try { if ($Host.UI.RawUI.WindowSize.Width -gt 0) { $width = $Host.UI.RawUI.WindowSize.Width } } catch {}
    $script:MenuWidth = [Math]::Max(20, [Math]::Min(76, $width - 8))
    $script:MenuIndent = ' ' * [Math]::Max(0, [int][Math]::Floor(($width - $script:MenuWidth) / 2))
}

function Write-MenuText([string]$Text, [ConsoleColor]$Color = 'Gray') {
    if (-not $script:MenuWidth) { Initialize-MenuLayout }
    # Wrap text within the menu column, including long file paths and prompts.
    foreach ($paragraph in ($Text -split "`r?`n")) {
        $remaining = $paragraph
        while ($remaining.Length -gt $script:MenuWidth) {
            $cut = $remaining.LastIndexOf(' ', $script:MenuWidth)
            if ($cut -lt 1) { $cut = $script:MenuWidth }
            Write-Host ($script:MenuIndent + $remaining.Substring(0, $cut)) -ForegroundColor $Color
            $remaining = $remaining.Substring($cut).TrimStart()
        }
        Write-Host ($script:MenuIndent + $remaining) -ForegroundColor $Color
    }
}

function Show-MenuOption([string]$Key, [string]$Label) {
    Write-MenuText ('  [{0}]  {1}' -f $Key, $Label) White
    Write-Host ''
}

function Show-Header([string]$Title) {
    Clear-MenuScreen
    Initialize-MenuLayout
    $script:MenuTitle = $Title
    Write-Host ''
    Write-Host ''
    Write-MenuText 'EXCHANGE ONLINE  /  ADMINISTRACAO' Cyan
    Write-MenuText ('=' * $script:MenuWidth) DarkCyan
    Write-Host ''
    Write-MenuText $Title White
    Write-MenuText 'Confirme o tenant antes de executar uma rotina.' DarkGray
    Write-Host ''
}

function Wait-Menu {
    $null = Read-MenuValue 'Enter para continuar'
}

function Initialize-Audit {
    $null = New-Item -ItemType Directory -Path $LogDirectory -Force -ErrorAction Stop
    $script:LogPath = Join-Path $LogDirectory ('mrm-{0}-{1}.jsonl' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), ([guid]::NewGuid().ToString('N').Substring(0,8)))
    Write-Audit 'Sessao' @{ Estado = 'Iniciada' }
}

function Write-Audit([string]$Event, $Data) {
    if (-not $script:LogPath) { throw 'Registro local indisponivel. Operacao bloqueada.' }
    [ordered]@{ Data = (Get-Date).ToString('o'); Evento = $Event; Dados = $Data } |
        ConvertTo-Json -Depth 7 -Compress |
        Add-Content -LiteralPath $script:LogPath -Encoding UTF8 -ErrorAction Stop
}

function Get-ExactPolicy([string]$Name) {
    $found = @(Get-RetentionPolicy -ErrorAction Stop | Where-Object { $_.Name -ieq $Name })
    if ($found.Count -ne 1) { throw "Politica exata nao encontrada: $Name" }
    return $found[0]
}

function Show-PolicyDetails($Policy) {
    $Policy | Format-List Name, IsDefault, RetentionPolicyTagLinks | Out-Host
    $tags = @(foreach ($link in $Policy.RetentionPolicyTagLinks) {
        Get-RetentionPolicyTag -Identity $link -ErrorAction Stop
    })
    $tags | Format-Table Name, Type, AgeLimitForRetention, RetentionAction, RetentionEnabled -AutoSize -Wrap | Out-Host
    if ($tags.Count -eq 0) { Write-Host 'AVISO: esta politica nao possui tags.' -ForegroundColor Yellow }
    if (@($tags | Where-Object { $_.RetentionAction -eq 'PermanentlyDelete' }).Count -gt 0) {
        Write-Host 'ATENCAO: esta politica existente possui tag de exclusao permanente!' -ForegroundColor Red
    }
}

function Show-QueryPolicies([string]$Search) {
    $policies = @(Get-RetentionPolicy -ErrorAction Stop |
        Where-Object { -not $Search -or $_.Name -ilike "*$Search*" } | Sort-Object Name)
    if ($policies.Count -eq 0) {
        Write-Host "Nenhuma politica encontrada para: $Search" -ForegroundColor Yellow
        Wait-Menu
        return
    }
    while ($true) {
        Write-Host "Politicas encontradas: $($policies.Count)" -ForegroundColor Cyan
        for ($i = 0; $i -lt $policies.Count; $i++) {
            $label = ''
            if (Test-ProtectedPolicy $policies[$i]) { $label = ' [PADRAO - SOMENTE CONSULTA]' }
            Show-MenuOption ($i + 1) ($policies[$i].Name + $label)
        }
        $choice = Read-MenuValue 'Numero da politica para consultar suas tags e contas'
        if ($null -eq $choice) { return }
        $number = 0
        if (-not [int]::TryParse($choice, [ref]$number) -or $number -lt 1 -or $number -gt $policies.Count) {
            Write-Host 'Opcao invalida.' -ForegroundColor Yellow
            continue
        }
        $policy = $policies[$number - 1]
        Show-PolicyDetails $policy
        Write-Host 'Consultando contas (somente leitura; pode demorar)...' -ForegroundColor Cyan
        $accounts = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop |
            Where-Object { [string]$_.RetentionPolicy -ieq [string]$policy.Name })
        $accounts | Format-Table DisplayName, PrimarySmtpAddress, RetentionPolicy -AutoSize -Wrap | Out-Host
        Write-Host "Total: $($accounts.Count) conta(s)."
        Wait-Menu
    }
}

function Show-Queries {
    while ($true) {
        Show-Header 'CONSULTAR RETENCAO'
        Show-MenuOption '1' 'Listar todas as politicas'
        Show-MenuOption '2' 'Buscar politicas por parte do nome (LIKE)'
        Show-MenuOption '3' 'Consultar uma conta'
        $choice = Read-MenuValue 'Opcao'
        if ($null -eq $choice) { return }
        try {
            switch ($choice) {
                '1' {
                    Show-QueryPolicies
                }
                '2' {
                    $name = Read-MenuValue 'Parte do nome da politica (ex.: NAORESPONDA; aceita *)'
                    if ($null -eq $name) { continue }
                    if (-not $name) { Write-Host 'Digite parte do nome ou use a opcao 1 para listar todas.'; Wait-Menu; continue }
                    Show-QueryPolicies $name
                }
                '3' {
                    $account = Read-MenuValue 'Alias ou e-mail da conta'
                    if ($null -eq $account) { continue }
                    Assert-SpecificIdentity $account
                    $mailbox = @(Get-Mailbox -Identity $account -ErrorAction Stop)
                    if ($mailbox.Count -ne 1) { throw 'A consulta nao retornou uma unica caixa.' }
                    $mailbox[0] | Format-List DisplayName, PrimarySmtpAddress, RetentionPolicy,
                        RetentionHoldEnabled, ElcProcessingDisabled, LitigationHoldEnabled, RetainDeletedItemsFor | Out-Host
                    if ($mailbox[0].RetentionPolicy) { Show-PolicyDetails (Get-ExactPolicy ([string]$mailbox[0].RetentionPolicy)) }
                    else { Write-Host 'Conta sem politica MRM associada.' -ForegroundColor Yellow }
                    Wait-Menu
                }
                default { Write-Host 'Opcao invalida.' -ForegroundColor Yellow; Wait-Menu }
            }
        } catch {
            if ($_.Exception.Message -in @('MRM_MENU_EXIT', 'MRM_MENU_HOME')) { throw }
            Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
            Wait-Menu
        }
    }
}

function Get-ExplicitTargets([string[]]$Identities) {
    if ($Identities.Count -eq 0) { throw 'Nenhuma conta informada.' }
    $seen = @{}
    foreach ($identity in $Identities) {
        Assert-SpecificIdentity $identity
        $matches = @(Get-Mailbox -Identity $identity -ErrorAction Stop)
        if ($matches.Count -ne 1) { throw "'$identity' nao identifica uma unica caixa." }
        $mailbox = $matches[0]
        if ([string]$mailbox.RecipientTypeDetails -notin @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox')) {
            throw "Tipo de caixa nao suportado: $identity ($($mailbox.RecipientTypeDetails))."
        }
        $guid = [guid]::Empty
        if (-not [guid]::TryParse([string]$mailbox.Guid, [ref]$guid) -or $guid -eq [guid]::Empty) {
            throw "Conta sem GUID valido: $identity"
        }
        $key = $guid.ToString()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            [pscustomobject]@{
                Guid = $key; Conta = [string]$mailbox.PrimarySmtpAddress
                Nome = [string]$mailbox.DisplayName; PoliticaAnterior = [string]$mailbox.RetentionPolicy
                RetentionHold = $mailbox.RetentionHoldEnabled; ElcDisabled = $mailbox.ElcProcessingDisabled
                LitigationHold = $mailbox.LitigationHoldEnabled; Recuperacao = [string]$mailbox.RetainDeletedItemsFor
            }
        }
    }
}

function Read-AccountsFile([string]$Path) {
    $Path = $Path.Trim().Trim('"')
    if ([IO.Path]::GetExtension($Path) -ine '.txt' -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw 'Informe o caminho de um arquivo .txt existente.'
    }
    $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop |
        ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -Unique)
    if ($lines.Count -eq 0) { throw 'O arquivo nao possui contas.' }
    foreach ($line in $lines) { Assert-SpecificIdentity $line }
    return $lines
}

function Invoke-ExplicitAssignment($Policy, [object[]]$Targets) {
    if ($Targets.Count -eq 0) { throw 'Nenhum destino explicito. Operacao bloqueada.' }
    $currentPolicy = Get-ExactPolicy ([string]$Policy.Name)
    if (Test-ProtectedPolicy $currentPolicy) { throw 'Politica padrao protegida. Aplicacao bloqueada.' }
    # Revalida a protecao no momento da escrita; nao usa enumeracao global como destino.
    foreach ($target in $Targets) {
        $result = [ordered]@{ Conta = $target.Conta; Guid = $target.Guid; Anterior = $target.PoliticaAnterior
            Solicitada = [string]$currentPolicy.Name; Atual = ''; Aplicacao = 'Nao executada'; Processamento = 'Nao solicitado'; Detalhe = '' }
        # Se o registro falhar, nao inicia a proxima alteracao.
        Write-Audit 'AplicacaoSolicitada' $result
        try {
            $freshPolicy = Get-ExactPolicy ([string]$currentPolicy.Name)
            if (Test-ProtectedPolicy $freshPolicy) { throw 'Politica tornou-se padrao. Operacao bloqueada.' }
            $before = Get-Mailbox -Identity $target.Guid -ErrorAction Stop
            if ([string]$before.RetentionPolicy -ine $target.PoliticaAnterior) {
                throw 'Politica da conta mudou desde a previa. Consulte novamente antes de aplicar.'
            }
            if ([string]$before.RetentionPolicy -ine [string]$currentPolicy.Name) {
                Set-Mailbox -Identity $target.Guid -RetentionPolicy $currentPolicy.Name -ErrorAction Stop
                $result.Aplicacao = 'Enviada; aguardando validacao'
            }
            $valid = $false
            for ($attempt = 1; $attempt -le 3; $attempt++) {
                $after = Get-Mailbox -Identity $target.Guid -ErrorAction Stop
                $result.Atual = [string]$after.RetentionPolicy
                if ($result.Atual -ieq [string]$currentPolicy.Name) { $valid = $true; break }
                if ($attempt -lt 3) { Start-Sleep -Seconds 2 }
            }
            if (-not $valid) { throw 'Atribuicao ainda nao confirmada pelo Exchange. Consulte a conta antes de repetir.' }
            $result.Aplicacao = 'Validada'
            if ($target.PoliticaAnterior -ieq [string]$currentPolicy.Name) { $result.Aplicacao = 'Ja associada; validada' }
            try {
                Start-ManagedFolderAssistant -Identity $target.Guid -ErrorAction Stop
                $result.Processamento = 'Solicitado (assincrono)'
            } catch {
                $result.Processamento = 'Falha na solicitacao'
                $result.Detalhe = $_.Exception.Message
            }
        } catch {
            $result.Detalhe = $_.Exception.Message
            $result.Aplicacao = 'Erro / nao validada'
        }
        # Exibe antes de registrar: se o disco falhar, o resultado remoto continua visivel.
        [pscustomobject]$result | Format-List | Out-Host
        Write-Audit 'ResultadoAplicacao' $result
        [pscustomobject]$result
    }
}

function Show-ApplyMenu($Policy) {
    :applyLoop while ($true) {
        Show-Header "APLICAR: $($Policy.Name)"
        if (Test-ProtectedPolicy (Get-ExactPolicy ([string]$Policy.Name))) { throw 'Politica padrao protegida.' }
        Show-MenuOption '1' 'Aplicar em UMA conta'
        Show-MenuOption '2' 'Aplicar somente nas contas de um arquivo .txt'
        $choice = Read-MenuValue 'Opcao'
        if ($null -eq $choice) { return }
        $identities = @()
        switch ($choice) {
            '1' {
                $account = Read-MenuValue 'Alias ou e-mail da conta'
                if ($null -eq $account) { continue applyLoop }
                $identities = @($account)
            }
            '2' {
                Write-Host 'Crie um .txt com UM alias ou e-mail por linha (sem virgulas ou ponto e virgula).' -ForegroundColor Yellow
                Write-Host 'Linhas vazias e comentarios iniciados por # sao ignorados. Duplicatas sao removidas.'
                Write-Host 'Tambem pode iniciar: .\Exchange-Admin.ps1 -ContasTxt "C:\Listas\contas.txt"'
                if ($script:ArquivoContas) { Write-Host "Enter usa: $script:ArquivoContas" }
                $path = Read-MenuValue 'Caminho do .txt'
                if ($null -eq $path) { continue applyLoop }
                if (-not $path) { $path = $script:ArquivoContas }
                if (-not $path) { throw 'Caminho nao informado.' }
                $identities = @(Read-AccountsFile $path)
            }
            default { Write-Host 'Opcao invalida.' -ForegroundColor Yellow; Wait-Menu; continue applyLoop }
        }
        Write-Host 'Validando TODOS os destinatarios antes de alterar qualquer conta...' -ForegroundColor Cyan
        $targets = @(Get-ExplicitTargets $identities)
        Show-PolicyDetails (Get-ExactPolicy ([string]$Policy.Name))
        $targets | Format-Table Conta, PoliticaAnterior, RetentionHold, ElcDisabled, LitigationHold, Recuperacao -Wrap -AutoSize | Out-Host
        Write-Host "Destinos unicos: $($targets.Count)" -ForegroundColor Cyan
        Write-Host 'ATENCAO: a nova MRM substitui a politica atual de cada conta acima, inclusive suas regras de arquivo.' -ForegroundColor Yellow
        Write-Host 'Mensagens antigas podem ser excluidas ao processar. Tags anteriores ja estampadas podem continuar atuando.' -ForegroundColor Yellow
        Write-Host 'Recuperacao depende do prazo da caixa. Holds/regras do Purview podem alterar o resultado.' -ForegroundColor Yellow
        if (@($targets | Where-Object { (Test-TrueValue $_.RetentionHold) -or (Test-TrueValue $_.ElcDisabled) }).Count) {
            Write-Host 'AVISO: ha caixas com MRM suspenso/desabilitado. O script nao altera esses controles.' -ForegroundColor Yellow
        }
        while ($true) {
            $confirm = Read-MenuValue "Aplicar aos $($targets.Count) destino(s) acima? [1] Aplicar | [2] Cancelar"
            if ($null -eq $confirm -or $confirm -eq '2') {
                Write-Host 'Aplicacao cancelada.' -ForegroundColor Yellow
                continue applyLoop
            }
            if ($confirm -eq '1') { break }
            Write-Host 'ERRO: resposta invalida. Digite 1 ou 2.' -ForegroundColor Red
        }
        $results = @(Invoke-ExplicitAssignment $Policy $targets)
        $results | Format-Table Conta, Aplicacao, Processamento -Wrap -AutoSize | Out-Host
        Write-Host 'A solicitacao ao assistente nao confirma conclusao do processamento das mensagens.' -ForegroundColor Yellow
        Write-Host "Registro: $script:LogPath"
        Wait-Menu
    }
}

function Show-PolicyCatalog {
    while ($true) {
        Show-Header 'CATALOGO DE POLITICAS'
        $policies = @(Get-RetentionPolicy -ErrorAction Stop | Sort-Object Name)
        if ($policies.Count -eq 0) { Write-Host 'Nenhuma politica encontrada.'; Wait-Menu; return }
        for ($i = 0; $i -lt $policies.Count; $i++) {
            $label = ''
            if (Test-ProtectedPolicy $policies[$i]) { $label = ' [PADRAO - BLOQUEADA]' }
            Show-MenuOption ($i + 1) ($policies[$i].Name + $label)
        }
        $choice = Read-MenuValue 'Numero da politica'
        if ($null -eq $choice) { return }
        $number = 0
        if (-not [int]::TryParse($choice, [ref]$number) -or $number -lt 1 -or $number -gt $policies.Count) {
            Write-Host 'Opcao invalida.'; Wait-Menu; continue
        }
        if (Test-ProtectedPolicy $policies[$number - 1]) {
            Write-Host 'Politica padrao: apenas consulta permitida.' -ForegroundColor Yellow; Wait-Menu; continue
        }
        Show-ApplyMenu $policies[$number - 1]
    }
}

function New-MrmConfiguration([string]$BaseName, [string[]]$Folders, [int]$Days) {
    if ($Days -lt 1 -or $Days -gt 24855) { throw 'Prazo permitido: 1 a 24855 dias.' }
    if ([string]::IsNullOrWhiteSpace($BaseName) -or $BaseName -match '[\x00-\x1f*?\[\]\\/]') { throw 'Nome vazio ou com caracteres nao permitidos.' }
    if ($Folders.Count -eq 0 -or @($Folders | Where-Object { $_ -notin @('Inbox','SentItems') }).Count -gt 0 -or
        @($Folders | Select-Object -Unique).Count -ne $Folders.Count) { throw 'Pastas invalidas ou repetidas.' }
    $policyName = "MRM - $BaseName - $Days dias"
    $specs = @(foreach ($folder in $Folders) {
        $label = if ($folder -eq 'Inbox') { 'Inbox' } else { 'Sent Items' }
        [pscustomobject]@{ Name = "$BaseName - $label $Days dias"; Type = $folder }
    })
    if ($policyName.Length -gt 64 -or @($specs | Where-Object { $_.Name.Length -gt 64 }).Count) { throw 'Nome longo: os nomes finais devem ter no maximo 64 caracteres. Abrevie o nome base.' }
    # Qualquer falha de leitura interrompe; nao confunde falta de permissao com nome disponivel.
    $existingPolicies = @(Get-RetentionPolicy -ErrorAction Stop)
    $existingTags = @(Get-RetentionPolicyTag -ErrorAction Stop)
    if (@($existingPolicies | Where-Object { $_.Name -ieq $policyName }).Count) { throw "Politica ja existe: $policyName. Use o catalogo; nada sera sobrescrito." }
    foreach ($spec in $specs) {
        if (@($existingTags | Where-Object { $_.Name -ieq $spec.Name }).Count) { throw "Tag ja existe: $($spec.Name). Nada sera sobrescrito; use outro nome ou revise no Exchange." }
    }
    Write-Host "Politica: $policyName" -ForegroundColor Cyan
    $specs | Format-Table Name, Type -AutoSize | Out-Host
    Write-Host "Prazo: $Days dias | Acao: DeleteAndAllowRecovery | Retencao habilitada"
    Write-Host 'Serao criadas apenas as tags acima, sem copiar tags da politica padrao.' -ForegroundColor Yellow
    while ($true) {
        $confirm = Read-MenuValue '[1] Criar tags e politica | [2] Cancelar'
        if ($null -eq $confirm -or $confirm -eq '2') { Write-Host 'Criacao cancelada.'; return }
        if ($confirm -eq '1') { break }
        Write-Host 'Resposta invalida. Digite 1 ou 2.' -ForegroundColor Yellow
    }
    $created = @()
    try {
        foreach ($spec in $specs) {
            Write-Audit 'CriarTagSolicitada' $spec
            $null = New-RetentionPolicyTag -Name $spec.Name -Type $spec.Type -AgeLimitForRetention $Days -RetentionAction DeleteAndAllowRecovery -RetentionEnabled $true -ErrorAction Stop
            $created += $spec.Name
            $tag = Get-RetentionPolicyTag -Identity $spec.Name -ErrorAction Stop
            if ([string]$tag.Type -ne $spec.Type -or [string]$tag.RetentionAction -ne 'DeleteAndAllowRecovery' -or
                -not (Test-TrueValue $tag.RetentionEnabled) -or (Get-RetentionDays $tag.AgeLimitForRetention) -ne $Days) { throw "Tag criada, mas nao validada: $($spec.Name)" }
            Write-Audit 'TagValidada' @{ Nome = $spec.Name }
        }
        Write-Audit 'CriarPoliticaSolicitada' @{ Nome = $policyName; Tags = @($specs.Name) }
        $null = New-RetentionPolicy -Name $policyName -RetentionPolicyTagLinks @($specs.Name) -ErrorAction Stop
        $created += $policyName
        $policy = Get-ExactPolicy $policyName
        if (Test-ProtectedPolicy $policy) { throw 'Politica identificada como padrao; fluxo interrompido.' }
        $linked = @(foreach ($link in $policy.RetentionPolicyTagLinks) { (Get-RetentionPolicyTag -Identity $link -ErrorAction Stop).Name })
        if ($linked.Count -ne $specs.Count -or @(Compare-Object @($specs.Name) $linked).Count) { throw 'Os vinculos da politica nao foram validados.' }
        Write-Audit 'PoliticaValidada' @{ Nome = $policyName; Tags = $linked }
        Write-Host 'SUCESSO: tags e politica criadas e validadas.' -ForegroundColor Green
        Show-PolicyDetails $policy
        return $policy
    } catch {
        Write-Host 'Criacao interrompida. NENHUMA conta foi alterada nesta etapa.' -ForegroundColor Red
        Write-Host "Objetos cuja criacao retornou sucesso: $($created -join '; ')" -ForegroundColor Yellow
        Write-Host 'Nao ha rollback por exclusao. Em falhas de rede, consulte tambem os nomes planejados antes de tentar novamente.' -ForegroundColor Yellow
        throw
    }
}

function Show-CreateMenu {
    :createLoop while ($true) {
        Show-Header 'CRIAR TAGS + POLITICA MRM'
        Show-MenuOption '1' 'Caixa de entrada (Inbox)'
        Show-MenuOption '2' 'Itens enviados (SentItems)'
        Show-MenuOption '3' 'As duas pastas (mesmo prazo)'
        $choice = Read-MenuValue 'Opcao'
        if ($null -eq $choice) { return }
        $folders = @()
        switch ($choice) {
            '1' { $folders = @('Inbox') }
            '2' { $folders = @('SentItems') }
            '3' { $folders = @('Inbox','SentItems') }
            default { Write-Host 'Opcao invalida.'; Wait-Menu; continue createLoop }
        }
        $name = Read-MenuValue 'Nome base (ex.: ENVIOEXTERNO; sem prefixo MRM e sem prazo)'
        if ($null -eq $name) { continue }
        $inputDays = Read-MenuValue 'Prazo em dias'
        if ($null -eq $inputDays) { continue }
        $days = 0
        if (-not [int]::TryParse($inputDays, [ref]$days)) { Write-Host 'Digite dias inteiros.'; Wait-Menu; continue }
        $policy = New-MrmConfiguration $name $folders $days
        if ($policy) { Wait-Menu; Show-ApplyMenu $policy }
    }
}


$script:AppRoot = $PSScriptRoot
foreach ($module in @('Common','Reports','Mailboxes','Groups','Menus')) {
    . (Join-Path $PSScriptRoot "modules\$module.ps1")
}
if ($MyInvocation.InvocationName -ne '.') { Start-ExchangeAdmin }
