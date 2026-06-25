# Guia de Implantação — Innova Excel Verifier

Este documento descreve como implantar o suplemento, realizar testes via sideload, hospedar os arquivos em servidores de produção e gerenciar configurações em produção.

---

## 1. Como fazer Sideload do Manifesto para Testes

O sideload permite carregar o suplemento localmente no Outlook de um usuário para desenvolvimento ou homologação.

### Testando no Outlook na Web (M365)
1. Certifique-se de que os arquivos do add-in compilados na pasta `dist` estão sendo servidos via HTTPS em `https://localhost:3000` (ex: rodando um servidor web local).
2. Acesse o [Outlook Web](https://outlook.office.com/).
3. Abra ou inicie a criação de um e-mail ("Novo E-mail").
4. Na barra de ferramentas de edição da mensagem, clique no ícone de três pontos (`...`) e selecione **Obter suplementos** ou **Aplicativos** > **Adicionar Aplicativos**.
5. No painel de Suplementos do Outlook, selecione **Meus suplementos**.
6. Role até a seção "Suplementos personalizados", clique em **Adicionar um suplemento personalizado** e escolha **Adicionar de um arquivo...**.
7. Selecione o arquivo `manifest.xml` localizado na raiz do seu projeto local.
8. Clique em **Instalar** (ignore avisos de certificados autoassinados em ambiente de desenvolvimento).
9. O suplemento será carregado e estará ativo para testes locais.

### Testando no Novo Outlook para Windows
Como o Novo Outlook para Windows é baseado na mesma engine web do Outlook Web, a instalação realizada no Outlook Web sincroniza automaticamente com o cliente desktop se você estiver logado com a mesma conta corporativa do Microsoft 365.

---

## 2. Publicação Piloto via Microsoft 365 Admin Center

Para distribuir o add-in em produção de forma centralizada para um grupo de usuários ou para toda a empresa:

1. Acesse o portal do [Microsoft 365 Admin Center](https://admin.microsoft.com/) usando uma conta de Administrador Global ou de Aplicativos.
2. No menu lateral, acesse **Configurações** > **Aplicativos integrados**.
3. Clique no botão **Implantar aplicativo complementar**.
4. Na tela de implantação, selecione **Carregar manifesto personalizado (.xml) a partir deste dispositivo** ou informe o link público do manifesto hospedado.
5. Selecione o arquivo `manifest.xml` definitivo (com os domínios de produção atualizados no lugar de `https://localhost:3000`).
6. Clique em **Avançar**.
7. Escolha quem terá acesso ao add-in:
   - **Toda a organização**
   - **Usuários/Grupos específicos** (recomendado para piloto comercial: ex. grupo de TI ou Financeiro).
   - **Apenas eu** (para testes individuais finais).
8. Selecione o método de implantação:
   - **Fixo (Padrão)**: O add-in é instalado automaticamente para os usuários e eles não podem removê-lo.
   - **Disponível**: Os usuários podem instalar o add-in manualmente através da loja corporativa.
9. Revise as permissões solicitadas (`ReadWriteItem`) e clique em **Concluir implantação**.
10. O processo de propagação corporativa pelo Microsoft 365 pode levar de **6 a 24 horas** para aparecer em todos os clientes Outlook dos usuários selecionados.

---

## 3. Hospedagem dos Arquivos Estáticos (HTTPS)

O add-in é 100% estático (HTML, CSS, JS, JSON, Imagens). Portanto, ele pode ser hospedado em qualquer servidor web que suporte HTTPS.

### Alternativas de Hospedagem:
* **IIS (Internet Information Services - Windows Server)**:
  * Copie o conteúdo da pasta `dist` (gerada após o comando `npm run build`) para a pasta do site no IIS (ex: `C:\inetpub\wwwroot\excel-verifier`).
  * Certifique-se de configurar um certificado SSL válido e assinado por uma Autoridade Certificadora (CA) confiável.
* **AWS S3 + CloudFront**:
  * Faça o upload dos arquivos da pasta `dist` em um bucket privado do Amazon S3.
  * Crie uma distribuição no AWS CloudFront apontando para o bucket do S3 com HTTPS habilitado.
* **Azure Blob Storage + CDN**:
  * Hospede os arquivos em um container de armazenamento de blobs configurado como site estático no Microsoft Azure.
  * Habilite o Azure CDN com HTTPS.
* **Qualquer outro servidor estático (Nginx, Apache, Cloudflare Pages, etc.)**.

---

## 4. Cuidados Importantes no Deploy de Produção

### Configurações de MIME Types
Se você hospedar em servidores IIS ou Nginx, certifique-se de que o servidor está configurado para servir arquivos `.json` com o tipo MIME correto:
* Tipo MIME: `application/json`
Caso contrário, a requisição do `config.json` em tempo de execução pode falhar com erro HTTP 404 ou 403.

### Alteração das URLs do Manifesto para Produção
Antes de gerar a build de produção e enviar o manifesto para o Admin Center, você **deve** atualizar todas as ocorrências de `https://localhost:3000` no arquivo `manifest.xml` para a URL HTTPS pública definitiva onde os arquivos estarão hospedados (ex: `https://verifier.aatb.com.br/`).

---

## 5. Como Alterar Configurações no `config.json`

O add-in consome as regras dinamicamente através do arquivo `config.json` hospedado no servidor. Isso significa que a equipe de TI pode alterar as regras sem a necessidade de reconstruir o código ou reimportar o manifesto no Admin Center.

### Exemplo de Configuração (`config.json`):
```json
{
  "companyDomain": "@aatb.com.br",
  "supportedExtensions": [".xlsx", ".xlsm"],
  "startColumn": "M",
  "language": "pt-BR",
  "validateAllRecipients": true,
  "validateHiddenSheets": true,
  "allowUserOverride": true,
  "ignoreZipFiles": true
}
```

### Para aplicar uma nova regra:
1. Abra o arquivo `config.json` localizado no diretório raiz da sua hospedagem HTTPS pública.
2. Edite os parâmetros conforme desejado (ex: alterar `"startColumn"` de `"M"` para `"N"` ou mudar `"validateHiddenSheets"` para `false`).
3. Salve o arquivo.
4. Na próxima vez que qualquer usuário tentar enviar um e-mail, o add-in buscará as novas regras em tempo de execução e as aplicará instantaneamente.
