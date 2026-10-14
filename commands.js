/*
 * FinFormat – ciclos de formatação para modelagem financeira
 * Sem macros: usa Office.js, então cada clique vira UM passo de Ctrl+Z.
 *
 * Para mudar cores ou formatos, edite apenas o bloco CONFIG abaixo.
 */

const CONFIG = {
  // Ciclo de cor da fonte: preto -> azul (input) -> verde (link entre abas) -> vermelho (link externo)
  fontColors: ["#000000", "#0000FF", "#008000", "#FF0000"],

  // Ciclo de preenchimento: null = sem preenchimento
  fillColors: [null, "#FFF2CC", "#F2F2F2"],

  // Ciclo de formato numérico (positivo ; negativo ; zero)
  // Códigos sempre em notação americana: vírgula = milhar, ponto = decimal
  numberFormats: [
    "#,##0;(#,##0);-",
    "#,##0.0;(#,##0.0);-",
    "0.0%;(0.0%);-",
    '0.0"x";(0.0"x");-'
  ]
};

// ---------- utilitários ----------

function norm(value) {
  return value == null ? "" : String(value).replace(/\s/g, "").toUpperCase();
}

// Próximo item da lista; se o atual não estiver nela, volta para o primeiro.
function nextItem(list, current) {
  const i = list.findIndex((item) => norm(item) === norm(current));
  return list[(i + 1) % list.length];
}

/*
 * Executa um ciclo:
 *  1. lê o formato da célula ativa (referência)
 *  2. calcula o próximo item do ciclo
 *  3. aplica em todas as áreas selecionadas (inclui seleções com Ctrl+clique)
 * mergeUndoGroup garante que tudo vire um único Ctrl+Z.
 */
async function runCycle(event, cycle) {
  try {
    await Excel.run({ mergeUndoGroup: true }, async (context) => {
      const selection = context.workbook.getSelectedRanges();
      selection.areas.load("items");
      const active = context.workbook.getActiveCell();
      cycle.load(active);
      await context.sync();

      const next = cycle.next(active);
      selection.areas.items.forEach((range) => cycle.apply(range, next));
      await context.sync();
    });
  } catch (error) {
    console.error("FinFormat:", error);
  } finally {
    if (event && event.completed) event.completed();
  }
}

// ---------- os 3 ciclos ----------

const fontCycleDef = {
  load: (cell) => cell.load("format/font/color"),
  next: (cell) => nextItem(CONFIG.fontColors, cell.format.font.color),
  apply: (range, color) => { range.format.font.color = color; }
};

const fillCycleDef = {
  load: (cell) => cell.load("format/fill/color"),
  next: (cell) => {
    let current = cell.format.fill.color;
    if (norm(current) === "#FFFFFF") current = null; // sem preenchimento
    return nextItem(CONFIG.fillColors, current);
  },
  apply: (range, color) => {
    if (color == null) range.format.fill.clear();
    else range.format.fill.color = color;
  }
};

const numberCycleDef = {
  load: (cell) => cell.load("numberFormat"),
  next: (cell) => nextItem(CONFIG.numberFormats, cell.numberFormat[0][0]),
  apply: (range, format) => { range.numberFormat = format; } // valor único vale para todas as células
};

function fontCycle(event)   { return runCycle(event, fontCycleDef); }
function fillCycle(event)   { return runCycle(event, fillCycleDef); }
function numberCycle(event) { return runCycle(event, numberCycleDef); }

// Registro dos botões da Faixa de Opções
Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("fontCycle", fontCycle);
    Office.actions.associate("fillCycle", fillCycle);
    Office.actions.associate("numberCycle", numberCycle);
  }
});

// Compatibilidade com versões mais antigas do Excel
if (typeof window !== "undefined") {
  window.fontCycle = fontCycle;
  window.fillCycle = fillCycle;
  window.numberCycle = numberCycle;
}
