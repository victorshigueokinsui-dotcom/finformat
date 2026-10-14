# FinFormat – instalação

Add-in do Excel com 3 ciclos (Cor da Fonte, Preenchimento, Formato Numérico), sem macros e com Ctrl+Z.

## 1. Publicar no GitHub (tudo pelo navegador)

1. Em github.com, clique em **New repository**, nome **finformat**, marque **Public** e crie.
2. Clique em **uploading an existing file** e arraste todo o conteúdo desta pasta
   (`manifest.xml`, `commands.html`, `commands.js`, `LEIA-ME.md` e a pasta `assets`). Clique em **Commit changes**.
3. Vá em **Settings → Pages**. Em *Branch*, escolha **main** e **/(root)** e clique em **Save**.
   Em 1–2 minutos o site fica em `https://victorshigueokinsui-dotcom.github.io/finformat/`.
4. Teste no navegador: `https://victorshigueokinsui-dotcom.github.io/finformat/assets/logo-32.png` deve mostrar o ícone.

## 2. Manifesto

O `manifest.xml` já aponta para `https://victorshigueokinsui-dotcom.github.io/finformat/`. Não precisa editar.

## 3. Instalar no Excel para Windows

1. Crie uma pasta, por exemplo `C:\FinFormat`, e coloque o `manifest.xml` nela.
2. Botão direito na pasta → **Propriedades → Compartilhamento → Compartilhar** → compartilhe com você mesmo.
   Anote o caminho de rede exibido, por exemplo `\\NOME-DO-PC\FinFormat`.
3. No Excel: **Arquivo → Opções → Central de Confiabilidade → Configurações da Central de Confiabilidade
   → Catálogos de Suplementos Confiáveis**.
4. Em *URL do Catálogo*, cole o caminho de rede, clique em **Adicionar Catálogo**, marque **Mostrar no Menu** e dê OK.
5. Feche e reabra o Excel.
6. **Página Inicial → Suplementos → Mais Suplementos → PASTA COMPARTILHADA** → selecione **FinFormat** → **Adicionar**.

A aba **FinFormat** aparece na Faixa de Opções.

## 4. Testar

Para cada botão: selecione células, clique algumas vezes e aperte **Ctrl+Z**. Cada clique deve voltar um passo.
Teste também uma seleção com Ctrl+clique e uma coluna inteira.

## 5. Mudar cores e formatos

Edite o bloco `CONFIG` no topo do `commands.js` pelo GitHub. Depois, no `commands.html`,
aumente o número em `commands.js?v=1` (para `v=2`, `v=3`...). Feche e reabra o Excel.
Se a mudança não aparecer, limpe o cache apagando o conteúdo de
`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` com o Excel fechado.

## Atalhos de teclado

- Ctrl+Shift+C: ciclo de cor da fonte
- Ctrl+Shift+V: ciclo de preenchimento (conflita com "colar valores" do Excel; na primeira vez o Excel pergunta qual usar)
- Ctrl+Shift+N: ciclo de formato numérico

Para mudar as teclas, edite `shortcuts.json`.

## Notas

- Separadores americanos (1,234.5): **Arquivo → Opções → Avançado**, desmarque *Usar separadores do sistema*
  e defina decimal `.` e milhar `,`.
- O ciclo usa a **célula ativa** como referência e aplica o próximo formato a toda a seleção.
- Exige Microsoft 365 atualizado.
