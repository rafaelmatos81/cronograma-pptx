# Cronograma para PowerPoint

Suplemento do PowerPoint (Mac, Windows e web) que gera Gantt e linha do tempo **editáveis**, desenhados com
formas nativas, a partir de Excel, CSV, texto colado do Excel ou XML do MS Project.

- Linha de base sob cada barra e coluna de desvio (dias corridos, atrasos em destaque)
- Caminho crítico em destaque, percentual executado dentro da barra, data de status
- Modo lista (hierárquico, como o MS Project) ou faixas (swimlanes)
- Vínculos entre tarefas (setas em cotovelo)
- Filtros: ramos da EDT, níveis, só marcos, só críticas, atraso mínimo, em andamento, nome, período, tarefas ocultas
- Paginação automática: as páginas extras viram slides logo depois, e os grupos-pai se repetem com "(cont.)"
- Título do slide automático (término previsto e diferença para a linha de base) ou escrito por você
- Paletas prontas ou cores próprias, e escolha da fonte
- **Edição depois de inserido**: selecione o cronograma no slide, clique em Editar, mude o que quiser e clique em
  Atualizar. A configuração e os dados ficam gravados no próprio .pptx.

Os arquivos de cronograma são lidos dentro do PowerPoint e não são enviados a nenhum servidor.

## Instalar no Mac (uso pessoal)

1. Baixe o [manifest.xml](https://rafaelmatos81.github.io/cronograma-pptx/manifest.xml).
2. Copie para a pasta de suplementos do PowerPoint:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   cp manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
   ```
3. Feche e abra o PowerPoint. Em **Página Inicial › Suplementos** (ou **Inserir › Suplementos › Meus Suplementos**),
   o **Cronograma** aparece na lista de suplementos de desenvolvedor. Depois de aberto uma vez, o botão fica na
   guia Página Inicial.

No Windows, use o mesmo manifest.xml por uma pasta compartilhada de catálogo confiável
(Arquivo › Opções › Central de Confiabilidade › Catálogos de Suplementos Confiáveis).

## Como usar

1. Selecione o slide onde o cronograma deve entrar (um layout "Somente título" funciona bem).
2. Em **Dados**, abra o arquivo. A aba e as colunas são reconhecidas sozinhas; confira em "Colunas reconhecidas".
   Arquivos .mpp: no MS Project, use Arquivo › Salvar como › XML.
3. Ajuste **Visão**, **Tarefas** e **Aparência** olhando a prévia.
4. Clique em **Inserir no slide**.

Para posicionar o cronograma num lugar específico, desenhe um retângulo no slide, selecione-o e use
**Aparência › Usar forma selecionada** antes de inserir.

## Desenvolvimento

```bash
npm install
npm run certificado   # uma vez: certificado https para localhost
npm run dev           # https://localhost:3000
npm test              # testes do núcleo
npm run manifesto     # regenera manifest.xml e manifest-dev.xml
```

Para testar a versão local no PowerPoint, copie `manifest-dev.xml` para a pasta `wef` (passo 2 acima).
Fora do PowerPoint, `http://localhost:3000/?exemplo` abre o painel com um cronograma de exemplo, em modo prévia.

Teste com um cronograma real (o arquivo não vai para o git):

```bash
CRONOGRAMA_TESTE=caminho/cronograma.xlsx CRONOGRAMA_ABA="Nome da aba" npm test
```

Estrutura:

```
src/nucleo/   leitura, seleção, escala, layout e prévia SVG (sem dependência do Office)
src/powerpoint.ts   desenho, agrupamento, gravação e atualização no PowerPoint (Office.js)
src/painel.ts       painel lateral
```

`npm run publicar` roda os testes, gera o build e publica no GitHub Pages (branch gh-pages).
