# Automação Contínua com Google Sheets e Planilha Padrão Salva

Este aprimoramento torna o preenchimento automático no Google Sheets 100% integrado à sua rotina diária: o link da sua planilha fica salvo permanentemente no aplicativo, permitindo que você apenas acesse a página, envie seus arquivos de manifesto (PDF ou imagem) e, após uma rápida revisão na tela, grave automaticamente todos os volumes nas rotas da planilha online.

> [!IMPORTANT]
> **Decisões confirmadas:**
> - **Planilha Padrão Salva**: O link e a aba da planilha online do Google Sheets serão salvos automaticamente no seu navegador (`localStorage`), conectando e carregando as rotas assim que o app for aberto, sem precisar colar o link toda vez.
> - **Fluxo com Revisão**: Ao carregar o manifesto (PDF ou imagem), o sistema extrai as lojas e exibe a tabela de conferência para conferir volumes, rotas e divergências antes de gravar online.
> - **Preenchimento Automático em 1 Clique**: Ao clicar em confirmar, o sistema executa a gravação atômica direta via API do Google Sheets nas células corretas de cada rota e recalcula os totais.

---

## 1. Visão Geral e Conceito Central

- **O que faz**: Transforma o site no hub central diário de expedição e conferência:
  1. A planilha online da sua operação já fica carregada e conectada por padrão.
  2. Você arrasta ou seleciona os manifestos do dia (PDFs de transferência ou fotos da carga).
  3. O sistema lê as lojas e volumes instantaneamente (via extração nativa de PDF ou OCR de imagem).
  4. Você faz uma checagem visual rápida (podendo adicionar novos manifestos ou ajustar números se necessário).
  5. Clica em "Preencher no Google Sheets" e a planilha online é preenchida na mesma hora, atualizando data, pedido, volumes e totais.
- **Público-alvo**: Conferentes, estoquistas e encarregados de expedição que alimentam planilhas de rotas diariamente.
- **Benefício principal**: Economia de dezenas de minutos diários de digitação manual e eliminação total de erros humanos de digitação de volumes.

---

## 2. Experiência do Usuário e Design Visual

### Fluxo Simplificado do Dia a Dia
1. **Acesso ao Site**:
   - O aplicativo já reconhece a planilha padrão configurada e seu login do Google.
   - Um badge elegante mostra: `✓ Conectado a: [Nome da sua Planilha] · [Aba Ativa]`.
   - Há um botão discreto de engrenagem para "Alterar planilha padrão" caso precise trocar no futuro.
2. **Envio do Manifesto**:
   - Área de drag-and-drop intuitiva: basta arrastar os PDFs de manifesto da pasta do computador diretamente para o site.
   - Leitura automática em tempo real com indicador de progresso (ex: "Lendo página 1 de 3...").
3. **Revisão e Conferência**:
   - Tabela com lojas encontradas, destacando em verde as lojas que estão na rota da planilha e em amarelo quaisquer lojas avulsas.
   - Edição rápida inline se precisar corrigir algum número.
4. **Gravação Instantânea**:
   - Botão de ação destacado: `Preencher no Google Sheets`.
   - Confirmação rápida e gravação atômica via Google Sheets API.
   - Notificação de sucesso com link direto para abrir a planilha preenchida no navegador.

### Diretrizes Visuais
- Visual corporativo e limpo, sem poluição visual.
- Área de dropzone destacada para facilitar o envio rápido de arquivos.
- Numerais tabulares claros para quantidades de volumes e códigos de filial.

---

## 3. Decisões Técnicas e Arquitetura

- **Persistência de Configuração**: Chaves de configuração salvas com segurança no `localStorage` (`gs_default_spreadsheet_id`, `gs_default_sheet_name`, `gs_default_url`), carregando automaticamente na inicialização.
- **Reconexão Silenciosa de Sessão**: Verificação automática de autenticação existente do Google via Firebase Auth ao abrir o site.
- **Gravação em Lote Otimizada**: Utilização do endpoint `values:batchUpdate` para preencher centenas de células em uma única chamada de rede de menos de 1 segundo.

---

## 4. Diagrama da Arquitetura do Fluxo Diário

```
┌────────────────────────────────────────────────────────────────────────┐
│                        Rotina Diária do Usuário                        │
├────────────────────────────────────────────────────────────────────────┤
│ 1. Abre o site ───► Planilha padrão Google Sheets conecta sozinha      │
│ 2. Envia o manifesto (arrasta ou clica no PDF/Imagem)                  │
│ 3. Sistema extrai lojas, volumes e cruza com a planilha                │
│ 4. Conferência visual das lojas identificadas                          │
│ 5. Clica em "Preencher no Google Sheets" ──► Planilha online gravada   │
└────────────────────────────────────────────────────────────────────────┘
```
