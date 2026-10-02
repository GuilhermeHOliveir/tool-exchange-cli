# Tool Exchange CLI

Ferramenta interativa em PowerShell para rotinas de administração do Exchange Online e Microsoft 365. Reúne consultas, relatórios e operações controladas em menus no terminal, com prévia, confirmação e registros locais.

## Instalação

Use Windows PowerShell 5.1 ou PowerShell 7. Tenha permissão no tenant para as rotinas que pretende executar e instale os módulos necessários:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Reports,Microsoft.Graph.Users.Actions -Scope CurrentUser
```

Baixe pelo botão **Code → Download ZIP** do GitHub e extraia o arquivo, ou clone o repositório:

```powershell
git clone https://github.com/GuilhermeHOliveir/tool-exchange-cli.git
cd tool-exchange-cli
```

Para abrir o menu, dê dois cliques em **Exchange-Admin.bat** na pasta extraída ou clonada. Mantenha o `.bat`, o `Exchange-Admin.ps1` e a pasta `modules/` juntos. O iniciador usa o Windows PowerShell 5.1, aceita argumentos opcionais do script e mantém a janela aberta em caso de erro para permitir a leitura da mensagem.

Para executar manualmente no PowerShell, use `.\Exchange-Admin.ps1` na pasta do projeto.

O programa solicita autenticação ao abrir uma rotina que precisa de Exchange ou Microsoft Graph. Ele não instala módulos nem salva credenciais. Se a política de execução da sua organização bloquear o arquivo, siga o procedimento interno para liberar scripts.

## Uso

Revise o tenant e os destinatários antes de confirmar alterações. No fluxo que usa CSV, marque `Processar=SIM` apenas nas linhas aprovadas e salve o arquivo antes de continuar.

Para configurar exceções do relatório de inatividade, copie `config.example.json` para `config.local.json` e edite apenas a cópia local. O programa grava resultados em `relatorio/` e registros em `logs/`; esses arquivos não são publicados no Git.

**Atenção:** o sufixo `.desativado` no endereço não bloqueia login nem recebimento de mensagens; o endereço antigo permanece como alias. A remoção opcional de licenças diretas pode afetar outros serviços Microsoft 365. Faça primeiro uma consulta ou simulação e valide em uma conta de teste.

Licenciado sob a [MIT License](LICENSE).
