# Assinatura 9Net no Novo Outlook — instalação manual

Este é o pacote enxuto de produção. Ele contém somente:

- os arquivos estáticos que devem ser enviados ao GitHub;
- o `manifest.xml` que será enviado ao Microsoft 365;
- este passo a passo.

Não há PowerShell, CMD, GitHub CLI, Node.js, compilação ou GitHub Actions.

## Endereços já configurados

- Repositório: `https://github.com/Suporte9Net/9net-assinatura-outlook`
- GitHub Pages: `https://suporte9net.github.io/9net-assinatura-outlook/`
- Redirect do Microsoft Entra: `brk-multihub://suporte9net.github.io`

Para não precisar alterar as URLs do `manifest.xml`, use exatamente o nome de repositório `9net-assinatura-outlook`.

---

# Parte 1 — criar o repositório no GitHub

1. Entre em `https://github.com/Suporte9Net`.
2. Clique no botão **New** para criar um repositório.
3. Em **Repository name**, coloque exatamente:

   `9net-assinatura-outlook`

4. Marque **Public**.
5. Não adicione README, `.gitignore` ou licença pela tela do GitHub.
6. Clique em **Create repository**.
7. No repositório vazio, clique em **uploading an existing file** ou em **Add file > Upload files**.
8. Extraia este ZIP no computador.
9. Abra a pasta extraída e envie **o conteúdo dela**, não a pasta externa inteira.

Na raiz do repositório precisam aparecer diretamente arquivos como:

- `index.html`
- `config.js`
- `autorun.html`
- `manifest.xml`
- `README.md`
- pasta `assets`

Se aparecer uma pasta `9net-outlook-pages-manual` e os arquivos estiverem dentro dela, o upload ficou um nível abaixo e as URLs não funcionarão.

10. Em **Commit changes**, coloque `Primeiro deploy da assinatura 9Net`.
11. Confirme o commit na branch `main`.

---

# Parte 2 — habilitar o GitHub Pages

1. Dentro do repositório, abra **Settings**.
2. No menu esquerdo, abra **Pages**.
3. Em **Build and deployment**, escolha:

   - **Source:** `Deploy from a branch`
   - **Branch:** `main`
   - **Folder:** `/(root)`

4. Clique em **Save**.
5. Aguarde o GitHub concluir a publicação.
6. Abra:

   `https://suporte9net.github.io/9net-assinatura-outlook/`

Deve aparecer a mensagem **GitHub Pages está respondendo**.

Documentação oficial: `https://docs.github.com/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site`

---

# Parte 3 — criar o aplicativo no Microsoft Entra

Esta parte é necessária para a assinatura consultar os dados do usuário no Microsoft 365.

1. Entre em `https://entra.microsoft.com` com uma conta administradora.
2. Abra **Identity > Applications > App registrations**.
3. Clique em **New registration**.
4. Preencha:

   - **Name:** `9Net Outlook Signature`
   - **Supported account types:** somente contas deste diretório organizacional, ou seja, single tenant.
   - **Redirect URI — Platform:** `Single-page application (SPA)`
   - **Redirect URI:** `brk-multihub://suporte9net.github.io`

5. Clique em **Register**.
6. Na tela **Overview**, copie estes dois valores:

   - **Application (client) ID**
   - **Directory (tenant) ID**

7. No menu do aplicativo, abra **API permissions**.
8. Clique em **Add a permission > Microsoft Graph > Delegated permissions**.
9. Pesquise e marque somente:

   `User.Read`

10. Clique em **Add permissions**.
11. Clique em **Grant admin consent for 9Net** e confirme.

O status de `User.Read` deve ficar verde. Não crie segredo de cliente; este suplemento não utiliza segredo.

Documentação oficial do NAA: `https://learn.microsoft.com/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in`

---

# Parte 4 — colocar Client ID e Tenant ID no GitHub

1. Volte ao repositório no GitHub.
2. Abra o arquivo `config.js`.
3. Clique no lápis **Edit this file**.
4. Localize:

   `clientId: "CLIENT_ID_PENDENTE"`

5. Troque somente `CLIENT_ID_PENDENTE` pelo **Application (client) ID** copiado do Entra.
6. Localize:

   `tenantId: "TENANT_ID_PENDENTE"`

7. Troque somente `TENANT_ID_PENDENTE` pelo **Directory (tenant) ID**.
8. Não apague as aspas, a vírgula ou o restante do arquivo.
9. Clique em **Commit changes**.
10. Aguarde o GitHub Pages publicar a alteração.
11. Abra este endereço no navegador:

   `https://suporte9net.github.io/9net-assinatura-outlook/config.js`

Confirme que não aparece mais a palavra `PENDENTE`.

Client ID e Tenant ID são identificadores públicos. Não coloque senhas ou segredos no repositório.

---

# Parte 5 — escolher quais usuários receberão a assinatura

A maneira mais simples de controlar os usuários é criar um grupo de segurança no Microsoft Entra.

1. No `https://entra.microsoft.com`, abra **Identity > Groups > All groups**.
2. Clique em **New group**.
3. Configure:

   - **Group type:** `Security`
   - **Group name:** `9NET_ASSINATURA_OUTLOOK_NOVO`
   - **Membership type:** `Assigned`

4. Em **Members**, adicione os usuários que devem receber a assinatura.
5. Crie o grupo.

Depois, para liberar ou remover a assinatura de alguém, basta alterar os membros desse grupo.

Se já existir um grupo do AD local sincronizado com o Microsoft Entra e ele aparecer no portal do Microsoft 365, ele também pode ser usado. O grupo precisa estar visível no ambiente Microsoft 365; o filtro de segurança de uma GPO local sozinho não distribui o suplemento do Novo Outlook.

---

# Parte 6 — instalar o suplemento para os usuários escolhidos

1. Entre em `https://admin.microsoft.com` com uma conta de **Administrador Global** ou **Administrador do Exchange**.
2. Clique em **Show all/Mostrar tudo**.
3. Abra **Settings/Configurações > Integrated apps/Aplicativos integrados**.
4. Abra **Add-ins/Suplementos** e clique em **Deploy Add-in/Implantar suplemento**.

   Dependendo da interface do tenant, a opção pode aparecer como **Upload custom apps/Carregar aplicativos personalizados**.

5. Escolha a opção de enviar um suplemento personalizado por arquivo de manifesto.
6. Selecione o arquivo `manifest.xml` deste pacote.
7. Clique em **Upload/Carregar**.
8. Na seleção de usuários, escolha **Specific users/groups/Usuários ou grupos específicos**.
9. Pesquise e selecione:

   `9NET_ASSINATURA_OUTLOOK_NOVO`

10. Conclua em **Deploy/Implantar**.

Para testar antes, você pode escolher **Just me/Somente eu**. Depois abra o suplemento em **Aplicativos integrados**, entre na seção **Users/Usuários** e troque a atribuição para o grupo.

A Microsoft informa que a exibição do suplemento pode levar até 72 horas, embora normalmente seja mais rápida. Feche e abra novamente o Novo Outlook durante o teste.

Documentação oficial: `https://learn.microsoft.com/microsoft-365/admin/manage/manage-deployment-of-add-ins`

---

# Parte 7 — testar no Novo Outlook

1. Entre no Novo Outlook com um usuário que pertença ao grupo.
2. Feche e abra o Outlook novamente.
3. Crie uma mensagem nova.
4. A assinatura deve ser inserida automaticamente.
5. Para testar manualmente, abra **Apps/Aplicativos > Assinatura 9Net**.
6. Clique em **Entrar e testar**.
7. Confira os dados encontrados e clique em **Aplicar nesta mensagem**.

## Campos consultados

| Assinatura | Microsoft 365 / Graph | Comportamento |
|---|---|---|
| Nome | `displayName` | Exibido no cabeçalho. |
| Cargo | `jobTitle` | Some quando vazio. |
| E-mail | `mail` | Usa `userPrincipalName` se `mail` estiver vazio. |
| Telefone | `businessPhones[0]` | Usa o telefone corporativo padrão se vazio. |
| Ramal | `faxNumber` | Se vazio, não aparecem `|` nem `Ramal`. |
| Celular | `mobilePhone` | A linha inteira some quando vazio. |
| Endereço | `streetAddress`, `city`, `state`, `postalCode`, `country` | Montado em linhas na coluna Endereço. |

---

# Outlook clássico

O `manifest.xml` contém uma proteção específica para o Outlook clássico. Nele, o suplemento encerra o evento sem inserir outra assinatura.

Resultado:

- **Novo Outlook:** assinatura hospedada no GitHub Pages e distribuída pelo Microsoft 365.
- **Outlook clássico:** continua usando a GPO atual.
- Não deve existir assinatura duplicada.

---

# Solução rápida de problemas

## GitHub Pages mostra 404

Confirme que `index.html` está na raiz do repositório e que Pages está configurado como `main` + `/(root)`.

## O painel informa que a configuração está pendente

Edite `config.js`, coloque os IDs corretos e aguarde o novo deploy do GitHub Pages.

## Acesso ao Microsoft Graph foi negado

Abra o aplicativo `9Net Outlook Signature` no Entra, entre em **API permissions** e confirme `User.Read` com consentimento administrativo verde.

## O suplemento não aparece para o usuário

Confirme que ele pertence ao grupo selecionado, reinicie o Outlook e considere o prazo de até 72 horas da implantação centralizada.

## O telefone ou endereço está errado

Corrija os campos do usuário no Microsoft 365/Entra. A assinatura apenas reproduz os campos cadastrados.

## Atualizações futuras

Enquanto o endereço do GitHub Pages continuar igual, alterações nos arquivos HTML, CSS, JavaScript ou imagens não exigem novo upload do `manifest.xml`. O GitHub Pages publicará a atualização e o suplemento passará a carregar os arquivos novos.
