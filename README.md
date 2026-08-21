# Assinatura corporativa 9Net — Novo Outlook / Outlook Web

Versão de correção automática: **2.0.2**.

## Objetivo

- inserir a assinatura automaticamente ao criar uma nova mensagem;
- manter o botão manual como contingência;
- consultar os dados completos no Microsoft Graph (`User.Read`);
- não depender do cache MSAL do painel para a assinatura aparecer;
- manter o Outlook clássico sob responsabilidade da GPO atual.

## Hospedagem usada

Base URL atual:

`https://9net-assinatura-outlook.pages.dev`

Arquivos importantes:

- `autorun.html` / `autorun.js`: runtime automático do Novo Outlook e Outlook Web;
- `classic-skip.js`: runtime do Outlook clássico que apenas encerra o evento;
- `taskpane.html` / `taskpane.js`: painel manual de teste;
- `signature.js`: autenticação, cache de perfil e montagem da assinatura;
- `manifest.xml`: manifesto já configurado com `OnNewMessageCompose`.

## Correção 2.0.2

O manifesto já possuía `OnNewMessageCompose` e seguia o formato oficial da Microsoft. O problema mais frágil estava no runtime automático: ele tentava obter o Microsoft Graph antes de inserir qualquer assinatura. O painel manual e o runtime automático são runtimes diferentes; portanto, um login que funciona no painel não garante que o cache de autenticação esteja disponível no evento em segundo plano.

Agora o fluxo é:

1. `OnNewMessageCompose` dispara;
2. o runtime aplica imediatamente o perfil salvo em `Office.context.roamingSettings`;
3. se ainda não houver perfil salvo, aplica nome/e-mail do Outlook + dados corporativos fallback;
4. tenta o Microsoft Graph silenciosamente por até 4,5 segundos;
5. se o Graph responder, substitui pela assinatura completa e salva o perfil para as próximas mensagens;
6. o evento é finalizado com `event.completed()` mesmo em caso de erro.

Ao usar **Entrar e testar** no painel, o perfil completo também é salvo para o modo automático.

## Microsoft Entra / NAA

O app registration deve possuir a permissão delegada:

`User.Read`

Para Nested App Authentication, a Redirect URI precisa ser do tipo **Single-page application (SPA)** e corresponder ao domínio real do suplemento:

`brk-multihub://9net-assinatura-outlook.pages.dev`

Não use o redirect antigo do GitHub Pages se a aplicação está hospedada no Cloudflare Pages.

Client ID e Tenant ID ficam em `config.js`. Não use segredo de cliente no frontend.

## Publicação da correção

Para o primeiro teste desta versão, publique no Cloudflare Pages pelo menos:

- `autorun.js`
- `signature.js`
- `taskpane.js`
- `config.js`
- `index.html`
- `support.html`

O `manifest.xml` desta pasta mantém o mesmo ID e a mesma configuração de `LaunchEvent`. Como o evento já foi distribuído e aparece no Outlook, você pode testar primeiro apenas atualizando os arquivos hospedados, sem remover o add-in do Microsoft 365.

Depois de publicar:

1. feche todas as abas do Outlook Web e o Novo Outlook;
2. abra novamente;
3. abra **Aplicativos > Assinatura 9Net** uma vez;
4. clique em **Entrar e testar** para gravar o perfil completo;
5. feche o painel;
6. crie outro **Novo email**;
7. a assinatura deve aparecer sem clicar em **Aplicar nesta mensagem**.

### Teste zero-touch

Depois do teste acima, teste com outro usuário que nunca abriu o painel. A assinatura deve aparecer automaticamente com fallback imediatamente e, se o Graph silencioso funcionar nesse runtime, será atualizada com os dados completos.

## Outlook clássico

O `manifest.xml` aponta o override JavaScript do runtime automático para `classic-skip.js`. Assim o Outlook clássico encerra o evento e continua usando apenas a assinatura distribuída pela GPO, evitando duplicidade.
