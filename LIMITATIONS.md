# Limitações Técnicas — Innova Excel Verifier

Este documento descreve as limitações e restrições conhecidas do suplemento Innova Excel Verifier decorrentes da plataforma Microsoft Office.js, APIs do Outlook e dependências de terceiros.

---

## 1. Limitações do Office.js e Smart Alerts

### Limite de Tempo de Execução (Timeout)
O Outlook impõe limites rígidos de tempo para a execução de suplementos baseados em evento (`OnMessageSend`).
* Se a validação demorar mais do que o limite permitido (tipicamente entre **5 e 30 segundos**, dependendo do host e da rede), o Outlook finalizará o script do add-in automaticamente.
* O comportamento do Outlook em caso de timeout depende da política administrativa, podendo bloquear o envio ou permitir o envio silenciosamente.

### Compatibilidade de Host e Cliente
* O recurso Smart Alerts (`OnMessageSend` com `PromptUser`) é suportado apenas no **Outlook na Web (M365)** e no **Novo Outlook para Windows**.
* Clientes legados (como Outlook 2013/2016/2019/2021 de licença perpétua MSI) não suportam event-based activation moderna e irão ignorar o add-in, permitindo o envio de e-mails sem validação.

### Comportamento Offline
* Por rodar em um runtime headless que necessita baixar dependências dinâmicas (como buscar o `config.json` ou carregar o próprio add-in), a validação não funcionará em modo completamente offline.
* Se não houver conectividade com a internet, o Outlook pode bloquear todos os envios de mensagens ou ignorar as validações dependendo de como as políticas do Microsoft 365 Admin Center estiverem configuradas.

---

## 2. Limitações de Leitura de Anexos no Outlook

### Arquivos na Nuvem (Cloud Attachments)
* Anexos compartilhados como links do OneDrive ou SharePoint (Web Reference Attachments) não são recuperados pelo método `getAttachmentContentAsync` como arquivos físicos.
* O Outlook retorna esses itens com o tipo `WebReference`. Eles **não são validados** pelo add-in no MVP e são ignorados para evitar falsos negativos ou falhas no script.

### Limite de Tamanho de Anexo
* O processamento de arquivos excessivamente grandes (ex. planilhas de dezenas de megabytes) pode consumir muita memória no runtime headless e estourar o limite de tempo (timeout).
* Para planilhas gigantes, recomenda-se que os usuários utilizem compartilhamento via links em vez de anexos físicos, pois o add-in é otimizado para arquivos de tamanho corporativo típico (até 10-15 MB).

---

## 3. Limitações de Leitura e Parsing do Excel (SheetJS)

### Arquivos Protegidos por Senha / Criptografados
* A biblioteca SheetJS não consegue descriptografar arquivos protegidos por senha em JavaScript puro no navegador sem a chave.
* Nesses casos, o add-in detectará a falha durante a leitura (`PASSWORD_PROTECTED`) e exibirá um alerta em português de que o arquivo não pôde ser validado por motivos de segurança, deixando a critério do usuário enviar mesmo assim.

### Formatos Suportados
* O MVP suporta estritamente extensões `.xlsx` e `.xlsm`.
* Não há suporte para formatos legados como `.xls` (formato binário antigo Excel 97-2003), planilhas abertas `.ods` (OpenDocument), arquivos de texto `.csv` ou arquivos compactados `.zip`. Arquivos `.zip` serão ignorados mesmo se contiverem arquivos Excel internamente.

### Abas Muito Ocultas (Very Hidden)
* Planilhas Excel possuem três níveis de visibilidade de abas: Visível, Oculta (Hidden) e Muito Oculta (Very Hidden - ocultada via código VBA e não listada no menu "Reexibir" padrão do Excel).
* A biblioteca SheetJS identifica e lê o conteúdo de abas muito ocultas normalmente, pois acessa a estrutura XML subjacente do arquivo `.xlsx`/`.xlsm`. Portanto, o add-in **consegue validar e alertar** sobre abas muito ocultas normalmente.

### Fórmulas com Resultados Vazios
* Células contendo fórmulas cujo resultado avaliado seja vazio (ex: `=""` ou `=SE(VERDADEIRO;"";"")`) são identificadas como conteúdo válido pelo add-in no MVP porque o objeto da célula retornado pela biblioteca mantém a propriedade da fórmula (`cell.f`) ativa.
* Células que possuem fórmulas que retornam erros (como `#DIV/0!`, `#N/D`) também contam como conteúdo ativo e serão validadas normalmente.
