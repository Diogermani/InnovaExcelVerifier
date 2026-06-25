# Matriz de Testes — Innova Excel Verifier

Este documento descreve os cenários de teste manuais e unitários necessários para garantir o correto funcionamento do add-in em ambiente de homologação e produção.

---

## Matriz de Testes Recomendada

| ID | Cenário de Teste | Arquivo de Entrada | Resultado Esperado | Valida Requisito |
| :--- | :--- | :--- | :--- | :--- |
| **CT-01** | Ausência de Anexos | Nenhum | E-mail enviado sem exibir alertas. | RF08 |
| **CT-02** | Anexos Não Excel | `documento.pdf`, `foto.png` | E-mail enviado sem exibir alertas. | RF01 |
| **CT-03** | Excel sem conteúdo após coluna L | `limpo.xlsx` (conteúdo em A-L) | E-mail enviado sem exibir alertas. | RF03, RF08 |
| **CT-04** | Excel com conteúdo na coluna M | `violacao_m.xlsx` (célula M5 = "dados") | Alerta exibido informando violação na coluna M. Permite "Enviar mesmo assim". | RF03, RF08 |
| **CT-05** | Excel com conteúdo na coluna Z | `violacao_z.xlsx` (célula Z12 = 42) | Alerta exibido informando violação na coluna Z. | RF03 |
| **CT-06** | Múltiplas abas limpas | `multiplas_limpas.xlsx` (3 abas) | E-mail enviado sem exibir alertas. | RF04 |
| **CT-07** | Múltiplas abas com apenas uma problemática | `uma_aba_ruim.xlsx` (Abas: "Plan1" [OK], "Dados" [Ruim]) | Alerta exibido listando apenas a aba "Dados" como problemática. | RF04, RF09 |
| **CT-08** | Múltiplas abas problemáticas | `abas_ruins.xlsx` (Abas: "Vendas" [Ruim], "RH" [Ruim]) | Alerta exibido listando ambas as abas ("Vendas" e "RH") como problemáticas. | RF04, RF09 |
| **CT-09** | Aba oculta problemática | `oculta_ruim.xlsx` (Aba oculta: "Segredos") | Alerta exibido listando a aba oculta "Segredos" como problemática. | RF04, RF06 |
| **CT-10** | Linha oculta problemática | `linha_oculta.xlsx` (Linha 10 oculta, M10 = "x") | Alerta exibido identificando a violação na linha oculta. | RF06 |
| **CT-11** | Coluna oculta problemática | `col_oculta.xlsx` (Coluna O oculta, O5 = "x") | Alerta exibido identificando a violação na coluna oculta. | RF06 |
| **CT-12** | Células com apenas formatação | `formatado_vazio.xlsx` (Aba com bordas e cor de fundo em M:Z) | E-mail enviado sem exibir alertas (formatação não conta como conteúdo). | RF05 |
| **CT-13** | Arquivo protegido por senha | `protegido.xlsx` (Criptografado) | Alerta em português indicando impossibilidade de validação devido à proteção por senha. | RF07 |
| **CT-14** | Múltiplos anexos, apenas um ruim | `limpo.xlsx` + `violacao_m.xlsx` | Alerta exibido listando apenas o arquivo `violacao_m.xlsx` e sua aba problemática. | RF09 |
| **CT-15** | Múltiplos anexos problemáticos | `violacao_m.xlsx` + `protegido.xlsx` | Alerta exibido detalhando a violação de `violacao_m.xlsx` e listando `protegido.xlsx` como ilegível. | RF09, RF07 |
| **CT-16** | Célula com Fórmula Vazia | `formula_vazia.xlsx` (Célula M5 possui `=SE(1=2; "Sim"; "")`) | Alerta exibido! Fórmulas devem contar como conteúdo mesmo que o resultado seja vazio. | RF05 |

---

## Roteiro de Execução Manual

Para validar cada um dos cenários da tabela acima:

1. Abra o Outlook Web ou o Novo Outlook.
2. Inicie a redação de uma nova mensagem.
3. Anexe os arquivos conforme o cenário de teste desejado.
4. Clique em **Enviar**.
5. Observe se o comportamento condiz com o esperado (passou direto ou abriu o diálogo de Smart Alert em português).
6. No caso de alertas:
   * Teste clicar em **Não Enviar** (o envio deve ser abortado e você deve retornar à edição do e-mail).
   * Teste clicar em **Enviar mesmo assim** (o e-mail deve ser disparado normalmente).
