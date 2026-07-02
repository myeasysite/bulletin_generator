import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ImageRun, VerticalAlign,
  ShadingType, FootnoteReferenceRun
} from "https://esm.sh/docx@8.5.0";

/* ---------------------------------------------------------------------
 * CONFIG — column definitions & default oblast lists, taken from the
 * REEFMC bulletin template. Edit here if the bulletin structure changes.
 * ------------------------------------------------------------------- */

const TABLE1_COLUMNS = [
  { key: "fires",        label: "Numbers of fires" },
  { key: "area",         label: "Area of fires, ha" },
  { key: "forests",      label: "Forests" },
  { key: "coniferous",   label: "Thereof coniferous" },
  { key: "agri",         label: "Agricultural lands" },
  { key: "other",        label: "Other natural landscapes" },
  { key: "settlements",  label: "Settlements" },
  { key: "conservation", label: "Area of fires in nature conservation zones¹" }
];

const TABLE2_COLUMNS = [
  { key: "area",       label: "Area of fires, ha" },
  { key: "forests",    label: "Forests" },
  { key: "coniferous", label: "Thereof coniferous" },
  { key: "agri",       label: "Agricultural lands" },
  { key: "other",      label: "Other natural landscapes" },
  { key: "settlements",label: "Settlements" }
];

const TABLE1_OBLASTS = ["Cherkasy","Chernihiv","Chernivtsi","Crimea","Dnipropetrovsk",
  "Donetsk","Ivano-Frankivsk","Kharkiv","Kherson","Khmelnytskyi","Kirovohrad","Kyiv",
  "Luhansk","Lviv","Mykolaiv","Odesa","Poltava","Rivne","Sumy","Ternopil","Vinnytsia",
  "Volyn","Zakarpattia","Zaporizhzhia","Zhytomyr"];

const TABLE1_COMBAT_OBLASTS = ["Chernihiv","Dnipropetrovsk","Donetsk","Kharkiv",
  "Kherson","Kyiv","Luhansk","Sumy","Zaporizhzhia"];

const TABLE2_OBLASTS = ["Cherkasy","Chernihiv","Dnipropetrovsk","Donetsk",
  "Ivano-Frankivsk","Kharkiv","Kherson","Khmelnytskyi","Kirovohrad","Kyiv","Luhansk",
  "Lviv","Mykolaiv","Odesa","Poltava","Rivne","Sumy","Volyn","Zakarpattia",
  "Zaporizhzhia","Zhytomyr"];

const TABLE2_COMBAT_OBLASTS = ["Chernihiv","Dnipropetrovsk","Donetsk","Kharkiv",
  "Kherson","Kyiv","Sumy","Zaporizhzhia"];

function emptyRow(oblast, columns) {
  const row = { oblast };
  columns.forEach(c => row[c.key] = 0);
  return row;
}

/* ---------------------------------------------------------------------
 * STATE
 * ------------------------------------------------------------------- */

const state = {
  table1: TABLE1_OBLASTS.map(o => emptyRow(o, TABLE1_COLUMNS)),
  table1Combat: TABLE1_COMBAT_OBLASTS.map(o => emptyRow(o, TABLE1_COLUMNS)),
  table2: TABLE2_OBLASTS.map(o => emptyRow(o, TABLE2_COLUMNS)),
  table2Combat: TABLE2_COMBAT_OBLASTS.map(o => emptyRow(o, TABLE2_COLUMNS)),
  map1File: null,
  map2File: null
};

/* ---------------------------------------------------------------------
 * UI RENDERING
 * ------------------------------------------------------------------- */

function renderDataTable(containerId, rows, columns) {
  const wrap = document.getElementById(containerId);
  const table = document.createElement("table");
  table.className = "data-table";

  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  trh.innerHTML = `<th>Область</th>` + columns.map(c => `<th>${c.label}</th>`).join("") + `<th></th>`;
  thead.appendChild(trh);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((row, idx) => {
    const tr = document.createElement("tr");

    const oblastTd = document.createElement("td");
    oblastTd.className = "oblast-cell";
    const oblastInput = document.createElement("input");
    oblastInput.type = "text";
    oblastInput.value = row.oblast;
    oblastInput.addEventListener("input", e => row.oblast = e.target.value);
    oblastTd.appendChild(oblastInput);
    tr.appendChild(oblastTd);

    columns.forEach(col => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.type = "number";
      input.value = row[col.key];
      input.addEventListener("input", e => row[col.key] = Number(e.target.value) || 0);
      td.appendChild(input);
      tr.appendChild(td);
    });

    const delTd = document.createElement("td");
    const delBtn = document.createElement("button");
    delBtn.className = "btn-del";
    delBtn.type = "button";
    delBtn.textContent = "✕";
    delBtn.addEventListener("click", () => {
      rows.splice(idx, 1);
      renderDataTable(containerId, rows, columns);
    });
    delTd.appendChild(delBtn);
    tr.appendChild(delTd);

    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  wrap.innerHTML = "";
  wrap.appendChild(table);
}

function renderAll() {
  renderDataTable("table1Wrap", state.table1, TABLE1_COLUMNS);
  renderDataTable("table1CombatWrap", state.table1Combat, TABLE1_COLUMNS);
  renderDataTable("table2Wrap", state.table2, TABLE2_COLUMNS);
  renderDataTable("table2CombatWrap", state.table2Combat, TABLE2_COLUMNS);
}

document.querySelectorAll("[data-add]").forEach(btn => {
  btn.addEventListener("click", () => {
    const key = btn.dataset.add;
    const columns = key.startsWith("table1") ? TABLE1_COLUMNS : TABLE2_COLUMNS;
    state[key].push(emptyRow("Нова область", columns));
    renderAll();
  });
});

/* Bulk paste: parses tab-separated (Excel/Google Sheets) or comma-separated
   text into rows. First column = oblast name, remaining columns map to the
   table's numeric fields in order. Replaces all rows in that table. */
function parsePastedBlock(text, columns) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  return lines.map(line => {
    const parts = line.includes("\t") ? line.split("\t") : line.split(/,|;/);
    const row = { oblast: (parts[0] || "").trim() };
    columns.forEach((c, i) => {
      const raw = (parts[i + 1] || "0").trim().replace(",", ".");
      row[c.key] = Number(raw) || 0;
    });
    return row;
  });
}

document.querySelectorAll("[data-apply]").forEach(btn => {
  btn.addEventListener("click", () => {
    const key = btn.dataset.apply;
    const columns = key.startsWith("table1") ? TABLE1_COLUMNS : TABLE2_COLUMNS;
    const textarea = document.getElementById(key + "Paste");
    const parsed = parsePastedBlock(textarea.value, columns);
    if (parsed.length === 0) return;
    state[key] = parsed;
    renderAll();
    textarea.value = "";
  });
});

function setupFileInput(inputId, previewId, stateKey) {
  document.getElementById(inputId).addEventListener("change", e => {
    const file = e.target.files[0];
    state[stateKey] = file || null;
    const preview = document.getElementById(previewId);
    if (file) {
      preview.src = URL.createObjectURL(file);
      preview.hidden = false;
    } else {
      preview.hidden = true;
    }
  });
}
setupFileInput("map1", "map1Preview", "map1File");
setupFileInput("map2", "map2Preview", "map2File");

renderAll();

/* ---------------------------------------------------------------------
 * DOCX GENERATION
 * ------------------------------------------------------------------- */

const FONT = "Times New Roman";
const HEADER_FONT = "Arial";
const THIN_BORDER = { style: BorderStyle.SINGLE, size: 2, color: "000000" };
const NONE_BORDER = { style: BorderStyle.NONE, size: 0, color: "auto" };
const CELL_BORDERS = { top: THIN_BORDER, bottom: THIN_BORDER, left: THIN_BORDER, right: THIN_BORDER };
const NO_BORDERS = { top: NONE_BORDER, bottom: NONE_BORDER, left: NONE_BORDER, right: NONE_BORDER };
const BOTTOM_BORDER_ONLY = { top: NONE_BORDER, bottom: THIN_BORDER, left: NONE_BORDER, right: NONE_BORDER };
const TOP_BORDER_ONLY = { top: THIN_BORDER, bottom: NONE_BORDER, left: NONE_BORDER, right: NONE_BORDER };
const HEADER_FILL = "E7E6E6";

function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: 22, ...opts });
}

function para(children, opts = {}) {
  return new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });
}

function sumRows(rows, columns) {
  const totals = {};
  columns.forEach(c => totals[c.key] = rows.reduce((s, r) => s + (Number(r[c.key]) || 0), 0));
  return totals;
}

function pct(part, whole) {
  if (!whole) return 0;
  return Math.round((part / whole) * 100);
}

// Usable page width in DXA (twips): A4 (11906) minus left+right margins (851 each)
const TABLE_WIDTH_DXA = 11906 - 851 - 851;

function computeColumnWidths(weights) {
  const sum = weights.reduce((a, b) => a + b, 0);
  return weights.map(w => Math.round((w / sum) * TABLE_WIDTH_DXA));
}

const COMBAT_FILL = "FCE4D6"; // light peach, matches original bulletin shading

function borderForLine(line) {
  if (line === "all") return CELL_BORDERS;
  if (line === "bottom") return BOTTOM_BORDER_ONLY;
  if (line === "top") return TOP_BORDER_ONLY;
  return NO_BORDERS;
}

function dataCell(text, opts = {}) {
  return new TableCell({
    borders: borderForLine(opts.borderLine),
    verticalAlign: VerticalAlign.CENTER,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shade ? { type: ShadingType.CLEAR, fill: COMBAT_FILL } : undefined,
    children: [para(run(String(text), { bold: opts.bold, italics: opts.italics, size: 20, font: opts.font || FONT }), { alignment: AlignmentType.CENTER })]
  });
}

function headerCell(text, opts = {}) {
  return new TableCell({
    borders: CELL_BORDERS,
    verticalAlign: VerticalAlign.CENTER,
    rowSpan: opts.rowSpan,
    columnSpan: opts.columnSpan,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: { type: ShadingType.CLEAR, fill: HEADER_FILL },
    children: [para(run(String(text), { bold: true, size: 17, font: HEADER_FONT }), { alignment: AlignmentType.CENTER })]
  });
}

/**
 * Builds a two-row merged header, matching the original bulletin layout:
 * Oblast | [leading cols, each rowSpan 2] | Distribution of fire area... (colSpan) | [trailing cols, rowSpan 2]
 *                                          | Forests | Coniferous | Agri | Other | Settlements
 */
function buildHeaderRows(columns, widths, leadingCount, trailingCount, distributionLabel) {
  const oblastW = widths[0];
  const colW = widths.slice(1);

  const leading = columns.slice(0, leadingCount);
  const leadingW = colW.slice(0, leadingCount);
  const distribution = columns.slice(leadingCount, columns.length - trailingCount);
  const distributionW = colW.slice(leadingCount, columns.length - trailingCount);
  const trailing = trailingCount > 0 ? columns.slice(columns.length - trailingCount) : [];
  const trailingW = trailingCount > 0 ? colW.slice(columns.length - trailingCount) : [];
  const distributionTotalW = distributionW.reduce((a, b) => a + b, 0);

  const row1 = [
    headerCell("Oblast", { rowSpan: 2, width: oblastW }),
    ...leading.map((c, i) => headerCell(c.label, { rowSpan: 2, width: leadingW[i] })),
    headerCell(distributionLabel, { columnSpan: distribution.length, width: distributionTotalW }),
    ...trailing.map((c, i) => headerCell(c.label, { rowSpan: 2, width: trailingW[i] }))
  ];
  const row2 = distribution.map((c, i) => headerCell(c.label, { width: distributionW[i] }));

  // Legend row: single-letter column codes (A, B, C, ...), as in the original bulletin
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  const legendRow = [oblastW, ...leadingW, ...distributionW, ...trailingW]
    .map((w, i) => dataCell(letters[i], { width: w, borderLine: "all", font: HEADER_FONT }));

  return [
    new TableRow({ tableHeader: true, children: row1 }),
    new TableRow({ tableHeader: true, children: row2 }),
    new TableRow({ tableHeader: true, children: legendRow })
  ];
}

function buildStatsTable(columns, mainRows, combatRows, combatLabel, layout) {
  const mainTotals = sumRows(mainRows, columns);
  const combatTotals = sumRows(combatRows, columns);

  const widths = computeColumnWidths(layout.weights);
  const rows = buildHeaderRows(columns, widths, layout.leadingCount, layout.trailingCount, layout.distributionLabel);

  const dataRow = (r, shade, borderLine) => new TableRow({
    children: [
      dataCell(r.oblast, { bold: true, width: widths[0], shade, borderLine }),
      ...columns.map((c, i) => dataCell(r[c.key], { width: widths[i + 1], shade, borderLine }))
    ]
  });
  const totalRow = (totals, shade, borderLine) => new TableRow({
    children: [
      dataCell("Total, ha", { bold: true, width: widths[0], shade, borderLine }),
      ...columns.map((c, i) => dataCell(totals[c.key], { bold: true, width: widths[i + 1], shade, borderLine }))
    ]
  });

  mainRows.forEach(r => rows.push(dataRow(r, false)));
  // Line after the first Total row
  rows.push(totalRow(mainTotals, false, "bottom"));

  rows.push(new TableRow({
    children: [
      new TableCell({
        borders: NO_BORDERS,
        columnSpan: columns.length + 1,
        children: [para(run(combatLabel, { italics: true, size: 18 }))]
      })
    ]
  }));

  combatRows.forEach(r => rows.push(dataRow(r, true)));
  // Line before the last Total row
  rows.push(totalRow(combatTotals, true, "top"));

  rows.push(new TableRow({
    children: [
      dataCell("%", { bold: true, width: widths[0], shade: true }),
      ...columns.map((c, i) => dataCell(pct(combatTotals[c.key], mainTotals[c.key]), { bold: true, width: widths[i + 1], shade: true }))
    ]
  }));

  return new Table({
    width: { size: TABLE_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: widths,
    rows
  });
}

async function fileToUint8(file) {
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

function imageTypeFromFile(file) {
  if (file.type === "image/png") return "png";
  if (file.type === "image/jpeg") return "jpg";
  return "png";
}

async function buildLogoImage() {
  const resp = await fetch("logo.png");
  const buf = await resp.arrayBuffer();
  return new ImageRun({
    data: new Uint8Array(buf),
    type: "png",
    transformation: { width: 66, height: 44 }
  });
}

async function buildMapImage(file, fallbackText) {
  if (!file) {
    return para(run(fallbackText, { italics: true, color: "999999" }));
  }
  const data = await fileToUint8(file);
  return para(new ImageRun({
    data,
    type: imageTypeFromFile(file),
    transformation: { width: 680, height: 481 }
  }), { alignment: AlignmentType.CENTER });
}

const EXPLANATORY_NOTE_1 =
  "In the research, we have used the network of nature conservation areas of European " +
  "importance, which was created to implement the provisions of the Berne Convention on " +
  "the Protection of Wild Flora and Fauna and Natural Habitats in Europe (Emerald Network) " +
  "(http://emerald.net.ua/).";

const EXPLANATORY_NOTE_2_T1 =
  "Monthly totals in Columns B and C include the number of fires and area burned all over " +
  "Ukraine (total of columns D, F, G and H).";

const EXPLANATORY_NOTE_2_T2 =
  "Monthly totals in Column B include the area burned of fires over Ukraine (total of " +
  "columns C, E, F and G).";

const EXPLANATORY_NOTE_3 =
  "Military Combat Zone: Permanent corridor spanning 30 km on each side of the LoC. The " +
  "corridor is determined at the beginning of active military combat and continues to be " +
  "included in the monitoring areas even after ending of active military combat operations. " +
  "Reason: Unexploded explosive ordnance remained in this area and continues to pose a " +
  "threat of ignition. Direct impact of military operations - 60 km along frontline with " +
  "both sides (30+30) of most intensive shelling and concentration of troops. Daily front " +
  "line coordinates were provided by ZOI Network for define fires that occurred in the zone " +
  "of direct impact (ZDI).";

const METHODOLOGY_PARAGRAPHS = [
  "Daily front line coordinates were provided by ZOI Network for define fires that occurred " +
  "in the zone of direct impact (ZDI). Distribution of burned land cover types within fire " +
  "perimeters were mapped using the Copernicus Dynamic Land Cover map at 100 m resolution " +
  "for 2019 (Copernicus Global Land Operations \"Vegetation and Energy\"), which provides a " +
  "detailed description of land cover as of 2019 with a resolution of 100 m.",
  "The network of nature conservation areas of European importance, which was created to " +
  "implement the provisions of the Berne Convention on the Protection of Wild Flora and " +
  "Fauna and Natural Habitats in Europe (Emerald Network) (http://emerald.net.ua/), was used " +
  "for the analysis of areas of nature conservation value that were affected by fires.",
  "The analysis was performed using a free geographic information system QGIS (3.2.1), " +
  "which is one of the most functional and convenient desktop geographic information systems."
];

function explanatoryNotes(noteTexts) {
  return noteTexts.map((text, i) =>
    para([run(`${i + 1} `, { superScript: true, size: 18 }), run(text, { size: 18, italics: true })],
      { spacing: { after: 100 } })
  );
}

function heading(text) {
  return para(run(text, { bold: true, size: 26 }), { spacing: { before: 200, after: 120 } });
}

async function generateDocx() {
  const bulletinNumber = document.getElementById("bulletinNumber").value.trim();
  const reportDate = document.getElementById("reportDate").value.trim();
  const periodFrom = document.getElementById("periodFrom").value.trim();
  const periodTo = document.getElementById("periodTo").value.trim();
  const periodLabel = `${periodFrom} to ${periodTo}`;

  const logo = await buildLogoImage();
  const map1 = await buildMapImage(state.map1File, "[ Карта 1 не завантажена ]");
  const map2 = await buildMapImage(state.map2File, "[ Карта 2 не завантажена ]");

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE }
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 40, type: WidthType.PERCENTAGE },
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            children: [
              para(run("Regional Eastern Europe Fire Monitoring Center", { bold: true })),
              para(run("(REEFMC)", { bold: true }))
            ]
          }),
          new TableCell({
            width: { size: 20, type: WidthType.PERCENTAGE },
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            children: [para(logo, { alignment: AlignmentType.CENTER })]
          }),
          new TableCell({
            width: { size: 40, type: WidthType.PERCENTAGE },
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            children: [
              para(run("Регіональний Східноєвропейський центр моніторингу пожеж", { bold: true, underline: {} }), { alignment: AlignmentType.RIGHT })
            ]
          })
        ]
      })
    ]
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, right: 851, bottom: 1134, left: 851 }
        }
      },
      children: [
        headerTable,
        para(run("Address: Kyiv, 03041, Dniprovska Naberezhna St. 7А, office 104.", { italics: true, underline: {}, size: 18 })),
        para(run("Tel. +38 067 2611682", { italics: true, underline: {}, size: 18 })),
        para([
          run("E-mail: ", { italics: true, underline: {}, size: 18 }),
          run("reefmc@nubip.edu.ua", { italics: true, underline: {}, size: 18, color: "0563C1" }),
          run("   URL: ", { italics: true, underline: {}, size: 18 }),
          run("https://nubip.edu.ua/node/9083", { italics: true, underline: {}, size: 18, color: "0563C1" })
        ], { border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "000000", space: 4 } } }),
        para(
          [run("Information bulletin", { bold: true, size: 24 }), new FootnoteReferenceRun(1), run(` № ${bulletinNumber}`, { bold: true, size: 24 })],
          { spacing: { before: 160 }, alignment: AlignmentType.CENTER }
        ),
        para(run(reportDate, { size: 22 }), { alignment: AlignmentType.CENTER }),
        para(run("FIRES ON THE TERRITORY OF UKRAINE", { bold: true, size: 26 }), { spacing: { before: 120 }, alignment: AlignmentType.CENTER }),
        para([run("(in ecosystems, agricultural lands and in cities", { size: 20 }), new FootnoteReferenceRun(2), run(")", { size: 20 })], { alignment: AlignmentType.CENTER }),
        para([run("For the period of ", { size: 22 }), run(periodLabel, { size: 22, underline: {} })], { spacing: { after: 160 }, alignment: AlignmentType.CENTER }),

        heading(`1. Map of active fires on the territory of Ukraine for the period of ${periodLabel}`),
        map1,

        heading(`Table 1. Statistics of fires on the territory of Ukraine for the period of ${periodLabel}`),
        buildStatsTable(TABLE1_COLUMNS, state.table1, state.table1Combat,
          "Including the Direct Impact of Military Combat Zone 60 km (30+30 km)³:",
          {
            leadingCount: 2, trailingCount: 1,
            distributionLabel: "Distribution of fire area by types of landscapes², ha",
            weights: [15, 8, 10, 9, 9, 10, 10, 8, 12]
          }),
        para(run("Explanatory Note:", { italics: true, bold: true, size: 18 }), { spacing: { before: 120 } }),
        ...explanatoryNotes([EXPLANATORY_NOTE_1, EXPLANATORY_NOTE_2_T1, EXPLANATORY_NOTE_3]),

        heading(`2. Map of fires in the nature conservation areas (Emerald Network) for the period of ${periodLabel}`),
        map2,

        heading(`Table 2. Statistics of fires in the areas of nature conservation (Emerald network¹) for the period of ${periodLabel}`),
        buildStatsTable(TABLE2_COLUMNS, state.table2, state.table2Combat,
          "Including the Direct Impact of Military Combat Zone 60 km (30+30 km)³:",
          {
            leadingCount: 1, trailingCount: 0,
            distributionLabel: "Distribution of fire area by types of landscapes², ha",
            weights: [20, 14, 15, 14, 15, 14, 13]
          }),
        para(run("Explanatory Note:", { italics: true, bold: true, size: 18 }), { spacing: { before: 120 } }),
        ...explanatoryNotes([EXPLANATORY_NOTE_1, EXPLANATORY_NOTE_2_T2, EXPLANATORY_NOTE_3]),

        para(run("_".repeat(90)), { spacing: { before: 200 } }),
        heading("3. Methodology for fire monitoring"),
        ...METHODOLOGY_PARAGRAPHS.map(t => para(run(t, { size: 20 }), { spacing: { after: 140 } }))
      ]
    }],
    footnotes: {
      1: {
        children: [para([
          run("The bulletin was prepared with the financial support of the Swiss organization «Zoï Environment Network» ", { size: 18 }),
          run("www.zoinet.org", { size: 18, color: "0563C1" })
        ])]
      },
      2: {
        children: [para(run("Methodology for fire monitoring is presented on the last page of this bulletin", { size: 18 }))]
      }
    }
  });

  return Packer.toBlob(doc);
}

document.getElementById("generateBtn").addEventListener("click", async () => {
  const btn = document.getElementById("generateBtn");
  const status = document.getElementById("status");
  btn.disabled = true;
  status.className = "status";
  status.textContent = "Формую документ…";
  try {
    const blob = await generateDocx();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const num = document.getElementById("bulletinNumber").value.trim() || "N";
    a.href = url;
    a.download = `REEFMC_bulletin_${num}.docx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    status.className = "status ok";
    status.textContent = "Готово! Файл завантажено.";
  } catch (err) {
    console.error(err);
    status.className = "status error";
    status.textContent = "Помилка генерації: " + err.message;
  } finally {
    btn.disabled = false;
  }
});
