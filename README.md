# Innova Excel Verifier — Outlook Add-in

Este é um suplemento (Add-in) corporativo para Microsoft Outlook desenvolvido em **Office.js** e **TypeScript** com o objetivo de validar arquivos Excel anexados (`.xlsx` e `.xlsm`) antes do envio de e-mails.

O add-in impede o vazamento de dados confidenciais verificando se há conteúdo preenchido a partir da coluna M (inclusive). Caso detecte conteúdo, um alerta em português é exibido ao usuário com a opção de enviar o e-mail mesmo assim (Smart Alerts: `PromptUser` mode) ou cancelar o envio para correção.

---

## Escopo do MVP

* **Detecção Automática**: Monitora o evento `OnMessageSend` no Outlook para Web e Novo Outlook para Windows.
* **Validação Inclusiva**: Analisa todas as células a partir da coluna M (colunas M, N, O... até o fim real da planilha).
* **Varredura Completa**: Varre todas as abas (visíveis, ocultas e muito ocultas).
* **Configuração Centralizada**: Configurações dinâmicas carregadas via arquivo `config.json` no servidor web (editável pela TI).
* **Sem Backend**: Execução 100% client-side (no próprio navegador/runtime do Outlook) visando privacidade e performance.
* **Tratamento de Erros**: Detecta arquivos protegidos por senha, corrompidos ou falhas de leitura, emitindo alertas amigáveis.

---

## Tecnologias Utilizadas

* **Linguagem**: TypeScript / JavaScript (ES6+)
* **APIs Office**: Office.js (Mailbox Requirement Set 1.10+)
* **Leitura de Excel**: [SheetJS (xlsx)](https://sheetjs.com/)
* **Bundler**: Webpack
* **Testes**: Jest e ts-jest

---

## Estrutura de Pastas

```
├── assets/                  # Logos e ícones do suplemento
├── src/
│   ├── config/              # Configurações padrão e serviço de configuração
│   ├── helpers/             # Validadores Excel e leitores de anexos
│   ├── launchevent/         # Handler de ativação baseada em evento (OnMessageSend)
│   └── taskpane/            # Interface gráfica simples (Task Pane)
├── manifest.xml             # Manifesto oficial do suplemento
├── package.json             # Dependências e scripts do projeto
├── tsconfig.json            # Configuração do TypeScript
└── webpack.config.js        # Configuração do empacotador Webpack
```

---

## Instalação e Execução Local

### 1. Instalar Dependências

Certifique-se de ter o Node.js instalado (v16+ recomendado) e execute:

```bash
npm install
```

### 2. Configurar HTTPS Local

O Office.js exige estritamente que todos os recursos do suplemento sejam servidos via **HTTPS**. Para desenvolvimento local, você precisará gerar um certificado SSL autoassinado confiável.

Uma forma simples de rodar um servidor local HTTPS com seus certificados é utilizando um servidor como `http-server` com SSL ou configurando o Webpack Dev Server para usar HTTPS.

Se preferir usar o `webpack-dev-server` no futuro ou um servidor simples, você pode gerar as chaves localmente:

1. Instale a ferramenta `mkcert` (se aplicável ao seu SO) ou gere certificados utilizando o OpenSSL:
   ```bash
   openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 -nodes
   ```
2. Salve os arquivos `key.pem` e `cert.pem` na raiz do projeto.
3. Configure seu servidor web para apontar na porta `3000` servindo a pasta `dist` via HTTPS utilizando esses certificados.

*(Mais detalhes sobre alternativas de servidores locais estão descritos no arquivo [DEPLOY.md](file:///c:/PROJETOS/Innova_ExcelVerifier/DEPLOY.md)).*

### 3. Compilar o Projeto

Para gerar os arquivos empacotados na pasta `dist` em modo de desenvolvimento contínuo (Watch):

```bash
npm run dev
```

Para gerar a build otimizada de produção:

```bash
npm run build
```

---

## Validação do Manifesto

Antes de fazer o upload do suplemento para o Outlook, é altamente recomendável validar a estrutura do arquivo `manifest.xml`. Você pode usar a ferramenta oficial da Microsoft:

```bash
npx office-addin-manifest validate manifest.xml
```

---

## Como Testar o Add-in

Para testar o comportamento do suplemento localmente no **Outlook na Web** ou **Novo Outlook para Windows**, siga os passos descritos no arquivo [DEPLOY.md](file:///c:/PROJETOS/Innova_ExcelVerifier/DEPLOY.md) na seção de **Sideload**.

Os cenários práticos de teste e amostras recomendadas estão descritos no arquivo [TESTS.md](file:///c:/PROJETOS/Innova_ExcelVerifier/TESTS.md).
As limitações técnicas documentadas da plataforma podem ser consultadas em [LIMITATIONS.md](file:///c:/PROJETOS/Innova_ExcelVerifier/LIMITATIONS.md).
