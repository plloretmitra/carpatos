'use strict';
const docx = require('docx');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, Table, TableRow, TableCell, WidthType,
  ShadingType, BorderStyle, PageBreak, ImageRun
} = docx;
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// ── PALETTE ────────────────────────────────────────────────────────
const C = {
  navy:       "1B2A4A",
  navyMid:    "2C3E6B",
  navyLight:  "3D5382",
  gold:       "C49A0A",
  goldLight:  "E8C040",
  white:      "FFFFFF",
  offWhite:   "F7F4EE",
  textDark:   "1A1A1A",
  textMid:    "444444",
  sidebar:    "EBF0F7",
  red:        "CC2936",
  green:      "2D6A4F",
  lightGold:  "FDF3D0",
  historyBg:  "EEF2F7",
  motoBg:     "FFF8EC",
};

// ── GENERIC HELPERS ────────────────────────────────────────────────
const NB = {
  top:    { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left:   { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right:  { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideH:{ style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideV:{ style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function empty(space = 0) {
  return new Paragraph({ children: [], spacing: { after: space } });
}

function pageBreakPara() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ── HISTORY CAPSULE ────────────────────────────────────────────────
// Devuelve la celda de la cápsula. Se monta dentro de dayBlock() como
// segunda fila, para que nunca se separe de la cabecera de su día.
function historyCapsuleCell(title, text) {
  return new TableCell({
      width: { size: 9026, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: C.historyBg },
      margins: { top: 120, bottom: 120, left: 300, right: 300 },
      borders: {
        left: { style: BorderStyle.SINGLE, color: C.navyMid, size: 20 },
        top: { style: BorderStyle.NONE, size: 0 },
        bottom: { style: BorderStyle.NONE, size: 0 },
        right: { style: BorderStyle.NONE, size: 0 },
      },
      children: [
        new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: `⏳  ${title}`, color: C.navyMid, bold: true, size: 21, font: "Calibri" })]
        }),
        new Paragraph({
          spacing: { after: 0 },
          children: [new TextRun({ text, size: 20, color: C.textMid, font: "Calibri", italics: true })]
        })
      ]
  });
}

// ── MOTO BOX ──────────────────────────────────────────────────────
function motoBox(text) {
  return new Table({
    columnWidths: [9026],
    width: { size: 9026, type: WidthType.DXA },
    borders: NB,
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9026, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: C.motoBg },
      margins: { top: 140, bottom: 140, left: 300, right: 300 },
      borders: {
        left: { style: BorderStyle.SINGLE, color: C.gold, size: 24 },
        top: { style: BorderStyle.SINGLE, color: C.gold, size: 6 },
        bottom: { style: BorderStyle.SINGLE, color: C.gold, size: 6 },
        right: { style: BorderStyle.NONE, size: 0 },
      },
      children: [
        new Paragraph({
          spacing: { after: 80 },
          children: [new TextRun({ text: "🏍  ALTERNATIVA: Hazlo en moto", color: "8B5E00", bold: true, size: 23, font: "Calibri" })]
        }),
        new Paragraph({
          spacing: { after: 0 },
          children: [new TextRun({ text, size: 21, color: C.textDark, font: "Calibri" })]
        })
      ]
    })]})],
  });
}

// ── BODY TEXT HELPERS ──────────────────────────────────────────────
// breakBefore lleva el salto de página en el propio encabezado. Un
// pageBreakPara() suelto es un párrafo más: si la sección anterior
// termina al ras del margen inferior, ese párrafo cae en la página
// siguiente y su salto deja esa página en blanco (pasaba en Word).
function sectionHead(text, breakBefore = false) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: breakBefore,
    spacing: { before: 400, after: 100 },
    border: { bottom: { color: C.gold, size: 10, style: BorderStyle.SINGLE, space: 4 } },
    children: [new TextRun({ text: text.toUpperCase(), color: C.navy, bold: true, size: 38, font: "Georgia" })]
  });
}

function subsectionHead(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 80 },
    children: [new TextRun({ text, color: C.navyMid, bold: true, size: 28, font: "Georgia" })]
  });
}

function body(text, bold = false, italics = false, color = C.textDark) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, bold, italics, size: 22, color, font: "Calibri" })]
  });
}

function bullet(text, bold = false) {
  return new Paragraph({
    spacing: { after: 80 },
    indent: { left: 440, hanging: 440 },
    children: [
      new TextRun({ text: "▸  ", color: C.gold, bold: true, size: 22 }),
      new TextRun({ text, size: 22, color: C.textDark, bold, font: "Calibri" })
    ]
  });
}

function highlight(text, bgColor = C.lightGold) {
  return new Table({
    columnWidths: [9026],
    width: { size: 9026, type: WidthType.DXA },
    borders: NB,
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9026, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: bgColor },
      margins: { top: 120, bottom: 120, left: 280, right: 280 },
      borders: {
        left: { style: BorderStyle.SINGLE, color: C.gold, size: 16 },
        top: { style: BorderStyle.NONE, size: 0 },
        bottom: { style: BorderStyle.NONE, size: 0 },
        right: { style: BorderStyle.NONE, size: 0 },
      },
      children: [new Paragraph({
        spacing: { after: 0 },
        children: [new TextRun({ text, size: 22, color: C.textDark, font: "Calibri", italics: true })]
      })]
    })]})],
  });
}

function momentDelDia(text) {
  return new Table({
    columnWidths: [9026],
    width: { size: 9026, type: WidthType.DXA },
    borders: NB,
    rows: [new TableRow({ children: [new TableCell({
      width: { size: 9026, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: "1B2A4A" },
      margins: { top: 160, bottom: 160, left: 320, right: 320 },
      children: [
        new Paragraph({
          spacing: { after: 40 },
          children: [new TextRun({ text: "✦  EL MOMENTO DEL DÍA", color: C.goldLight, bold: true, size: 17, font: "Calibri", characterSpacing: 40 })]
        }),
        new Paragraph({
          spacing: { after: 0 },
          children: [new TextRun({ text: text, color: "B0C4DE", size: 19, font: "Georgia", italics: true })]
        }),
      ]
    })]})],
  });
}

function infoBox(title, lines, titleBg = C.navyMid, bodyBg = C.sidebar) {
  const titleRow = new TableRow({ children: [new TableCell({
    width: { size: 9026, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: titleBg },
    margins: { top: 120, bottom: 120, left: 240, right: 240 },
    children: [new Paragraph({
      spacing: { after: 0 },
      children: [new TextRun({ text: title, color: C.white, bold: true, size: 23, font: "Calibri" })]
    })]
  })]});
  const bodyRows = lines.map(line => new TableRow({ children: [new TableCell({
    width: { size: 9026, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: bodyBg },
    margins: { top: 80, bottom: 80, left: 240, right: 240 },
    children: [new Paragraph({
      spacing: { after: 0 },
      indent: { left: 240, hanging: 240 },
      children: [
        new TextRun({ text: "▪  ", color: C.navyMid, bold: true, size: 21 }),
        new TextRun({ text: line, size: 21, color: C.textDark, font: "Calibri" })
      ]
    })]
  })]}));
  return new Table({ columnWidths: [9026], width: { size: 9026, type: WidthType.DXA }, borders: NB, rows: [titleRow, ...bodyRows] });
}

// Devuelve la celda de cabecera. La monta dayBlock() como primera fila.
function dayHeaderCell(num, title, subtitle, km = null, regions = null) {
  const main = `DÍA ${num}`;
  const sub  = km ? `  ·  ${km}` : "";
  const regionPara = regions ? [new Paragraph({
    spacing: { after: 0 },
    keepNext: true,
    children: [new TextRun({ text: regions, color: "7B9BBF", size: 17, font: "Calibri", italics: true })]
  })] : [];
  return new TableCell({
      width: { size: 9026, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: C.navy },
      margins: { top: 120, bottom: 120, left: 280, right: 280 },
      children: [
        new Paragraph({
          spacing: { after: 40 },
          keepNext: true,
          children: [
            new TextRun({ text: main, color: C.goldLight, bold: true, size: 20, font: "Calibri" }),
            new TextRun({ text: sub, color: "7B90B0", size: 19, font: "Calibri" })
          ]
        }),
        new Paragraph({
          spacing: { after: 0 },
          keepNext: true,
          children: [new TextRun({ text: title, color: C.white, bold: true, size: 28, font: "Georgia" })]
        }),
        new Paragraph({
          spacing: { after: regions ? 16 : 0 },
          keepNext: true,
          children: [new TextRun({ text: subtitle, color: "B0C4DE", size: 20, font: "Calibri", italics: true })]
        }),
        ...regionPara,
      ]
  });
}

// Salto de página antes de un día. Va en un párrafo propio porque una tabla
// no admite pageBreakBefore. El día 1 no lo lleva: viene justo detrás del
// encabezado de la sección del itinerario.
function dayPageBreak() {
  return new Paragraph({ pageBreakBefore: true, spacing: { after: 0 } });
}

// ── DAY BLOCK ──────────────────────────────────────────────────────
// Cabecera de día + cápsula histórica en UNA sola tabla de dos filas,
// ambas con cantSplit. Al ser una única tabla, ningún renderizador
// puede dejar la cabecera huérfana al pie de página ni partir la
// cápsula: keepNext no lo conseguía ni en Word ni en LibreOffice.
function dayBlock(num, title, subtitle, km, regions, capTitle, capText) {
  return new Table({
    columnWidths: [9026],
    width: { size: 9026, type: WidthType.DXA },
    borders: NB,
    rows: [
      new TableRow({ cantSplit: true, children: [dayHeaderCell(num, title, subtitle, km, regions)] }),
      new TableRow({ cantSplit: true, children: [historyCapsuleCell(capTitle, capText)] }),
    ],
  });
}

function twoColTable(colA, colB) {
  const W1 = 4400, W2 = 4426, GAP = 200;
  return new Table({
    columnWidths: [W1, GAP, W2],
    width: { size: 9026, type: WidthType.DXA },
    borders: NB,
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: W1, type: WidthType.DXA }, borders: NB, children: colA }),
      new TableCell({ width: { size: GAP, type: WidthType.DXA }, borders: NB, children: [empty()] }),
      new TableCell({ width: { size: W2, type: WidthType.DXA }, borders: NB, children: colB }),
    ]})]
  });
}

function budgetRow(concept, perDay, total, isEven) {
  const bg = isEven ? C.offWhite : C.white;
  function cell(text, w) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: bg },
      margins: { top: 80, bottom: 80, left: 180, right: 180 },
      children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text, size: 21, color: C.textDark, font: "Calibri" })] })]
    });
  }
  return new TableRow({ children: [cell(concept,5200), cell(perDay,1900), cell(total,1926)] });
}

function budgetTable(rows) {
  const headerRow = new TableRow({ children: [
    new TableCell({ width:{size:5200,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:C.navy}, margins:{top:120,bottom:120,left:180,right:180}, children:[new Paragraph({spacing:{after:0},children:[new TextRun({text:"Concepto",color:C.white,bold:true,size:22,font:"Calibri"})]})] }),
    new TableCell({ width:{size:1900,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:C.navy}, margins:{top:120,bottom:120,left:180,right:180}, children:[new Paragraph({spacing:{after:0},children:[new TextRun({text:"Coste/día",color:C.white,bold:true,size:22,font:"Calibri"})]})] }),
    new TableCell({ width:{size:1926,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:C.navy}, margins:{top:120,bottom:120,left:180,right:180}, children:[new Paragraph({spacing:{after:0},children:[new TextRun({text:"Total (12 días)",color:C.white,bold:true,size:22,font:"Calibri"})]})] }),
  ]});
  const totalRow = new TableRow({ children: [
    new TableCell({ width:{size:5200,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:C.navyLight}, margins:{top:120,bottom:120,left:180,right:180}, children:[new Paragraph({spacing:{after:0},children:[new TextRun({text:"TOTAL ESTIMADO (sin vuelo, sin moto opcional)",color:C.white,bold:true,size:22,font:"Calibri"})]})] }),
    new TableCell({ width:{size:1900,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:C.navyLight}, margins:{top:120,bottom:120,left:180,right:180}, children:[new Paragraph({spacing:{after:0},children:[new TextRun({text:"~113 €",color:C.goldLight,bold:true,size:22})]})] }),
    new TableCell({ width:{size:1926,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:C.navyLight}, margins:{top:120,bottom:120,left:180,right:180}, children:[new Paragraph({spacing:{after:0},children:[new TextRun({text:"~1.477-2.105 €",color:C.goldLight,bold:true,size:22})]})] }),
  ]});
  return new Table({
    columnWidths: [5200, 1900, 1926],
    width: { size: 9026, type: WidthType.DXA },
    rows: [headerRow, ...rows, totalRow],
    borders: {
      top:    { style: BorderStyle.SINGLE, color: C.navy, size: 4 },
      bottom: { style: BorderStyle.SINGLE, color: C.navy, size: 4 },
      left:   { style: BorderStyle.NONE,   size: 0 },
      right:  { style: BorderStyle.NONE,   size: 0 },
      insideH:{ style: BorderStyle.SINGLE, color: "CCCCCC", size: 2 },
      insideV:{ style: BorderStyle.NONE,   size: 0 },
    }
  });
}

// ── COVER PAGE ─────────────────────────────────────────────────────
function buildCover() {
  const illustrationData = fs.readFileSync(path.join(__dirname, '0b90d467-Gemini_Generated_Image_z9kjhez9kjhez9kj.png'));

  return [
    new Table({
      columnWidths: [9026],
      width: { size: 9026, type: WidthType.DXA },
      borders: NB,
      rows: [new TableRow({ children: [new TableCell({
        width: { size: 9026, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: C.gold },
        margins: { top: 80, bottom: 80, left: 280, right: 280 },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { after: 0 },
            children: [new TextRun({ text: "G U Í A  D E  V I A J E", color: C.navy, bold: true, size: 18, font: "Calibri", characterSpacing: 60 })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { after: 0 },
            children: [new TextRun({ text: "E D I C I Ó N  E S P E C I A L  R U M A N Í A", color: C.navy, bold: true, size: 16, font: "Calibri", characterSpacing: 50 })]
          }),
        ]
      })]})],
    }),

    new Table({
      columnWidths: [9026],
      width: { size: 9026, type: WidthType.DXA },
      borders: NB,
      rows: [new TableRow({ children: [new TableCell({
        width: { size: 9026, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: C.navy },
        margins: { top: 0, bottom: 0, left: 0, right: 0 },
        children: [
          empty(300),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
            children: [new TextRun({ text: "★  ★  ★", color: C.gold, size: 30 })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 40, after: 20 },
            children: [new TextRun({ text: "El Secreto de los Cárpatos", color: C.white, bold: true, size: 80, font: "Georgia" })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 20, after: 20 },
            children: [new TextRun({ text: "──────── ♦ ────────", color: C.gold, size: 28 })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 20, after: 40 },
            children: [new TextRun({ text: "12 días en coche por la Rumanía que nadie te ha contado", color: C.goldLight, size: 32, font: "Georgia", italics: true })]
          }),
          empty(200),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
            children: [new ImageRun({ data: illustrationData, transformation: { width: 560, height: 420 }, type: "png" })]
          }),
          empty(240),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
            children: [new TextRun({ text: "MUNTENIA  ·  OLTENIA  ·  TRANSILVANIA  ·  MARAMUREȘ  ·  MOLDOVA  ·  BUCOVINA", color: C.goldLight, size: 17, font: "Calibri", characterSpacing: 22 })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
            children: [new TextRun({ text: "──────────────────────────────", color: C.gold, size: 20 })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
            children: [new TextRun({ text: "12 Días  ·  Julio 2027", color: "B0BEC5", size: 21, font: "Calibri" })]
          }),
          empty(200),
        ]
      })]})],
    }),

    new Table({
      columnWidths: [3009, 3009, 3008],
      width: { size: 9026, type: WidthType.DXA },
      borders: NB,
      rows: [new TableRow({ children: [
        new TableCell({ width:{size:3009,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:"002B7F"}, margins:{top:60,bottom:60,left:0,right:0}, children:[empty(0)] }),
        new TableCell({ width:{size:3009,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:"FCD116"}, margins:{top:60,bottom:60,left:0,right:0}, children:[empty(0)] }),
        new TableCell({ width:{size:3008,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:"CE1126"}, margins:{top:60,bottom:60,left:0,right:0}, children:[empty(0)] }),
      ]})],
    }),
  ];
}

// ── CLIMATE TABLE ───────────────────────────────────────────────────
function climateTable() {
  const headers = ["Zona", "Temp. media", "Mínima", "Máxima", "Lluvia (mm)", "Sol/día"];
  const colWidths = [2200, 1400, 1100, 1100, 1626, 1600];
  const rows = [
    ["Bucarest", "26 °C", "18 °C", "34 °C", "~40 mm", "~10 h"],
    ["Transilvania (ciudades)", "20 °C", "12 °C", "28 °C", "~70 mm", "~9 h"],
    ["Alta montaña (+2.000 m)", "10 °C", "4 °C", "18 °C", "~80 mm", "~8 h"],
    ["Maramureș", "22 °C", "13 °C", "30 °C", "~80 mm", "~9 h"],
    ["Bucovina", "21 °C", "12 °C", "29 °C", "~90 mm", "~8 h"],
  ];

  const headerRow = new TableRow({ children: headers.map((h, i) =>
    new TableCell({
      width: { size: colWidths[i], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: C.navy },
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: h, color: C.white, bold: true, size: 20, font: "Calibri" })] })]
    })
  )});

  const dataRows = rows.map((r, i) => new TableRow({ children: r.map((cell, j) =>
    new TableCell({
      width: { size: colWidths[j], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: i % 2 === 0 ? C.offWhite : C.white },
      margins: { top: 70, bottom: 70, left: 160, right: 160 },
      children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: cell, size: 20, font: "Calibri", color: C.textDark, bold: j === 0 })] })]
    })
  )}));

  return new Table({
    columnWidths: colWidths,
    width: { size: 9026, type: WidthType.DXA },
    rows: [headerRow, ...dataRows],
    borders: {
      top:     { style: BorderStyle.SINGLE, color: C.navy, size: 4 },
      bottom:  { style: BorderStyle.SINGLE, color: C.navy, size: 4 },
      left:    { style: BorderStyle.NONE, size: 0 },
      right:   { style: BorderStyle.NONE, size: 0 },
      insideH: { style: BorderStyle.SINGLE, color: "CCCCCC", size: 2 },
      insideV: { style: BorderStyle.SINGLE, color: "DDDDDD", size: 2 },
    }
  });
}

// ── DOCUMENT ASSEMBLY ───────────────────────────────────────────────
const sections = [
  {
    properties: {},
    children: [
      ...buildCover(),
      pageBreakPara(),

      // ═══════════════════════════════════════════════════════
      // LO ESENCIAL
      // ═══════════════════════════════════════════════════════
      sectionHead("Lo Esencial"),
      body("Antes de entrar en los detalles del viaje, aquí van los datos que conviene tener claros desde el primer momento. El resto de secciones de esta guía los desarrollan en profundidad."),
      empty(100),
      infoBox("10 cosas imprescindibles que saber sobre Rumanía", [
        "Sin visado para ciudadanos de la UE — el DNI español es suficiente",
        "Moneda: Leu rumano (RON). 1 € ≈ 5 RON. En zonas rurales el efectivo es imprescindible",
        "Conducción por la derecha — igual que en España. Luces de cruce obligatorias de día",
        "Rovinieta obligatoria: peaje electrónico para circular por carreteras nacionales (~7 € / 30 días)",
        "Emergencias: número 112 (funciona en rumano e inglés)",
        "Huso horario: UTC+3 en verano (1 hora más que en España peninsular en julio)",
        "Presupuesto orientativo: ~113 € por día durante 13 días (alojamiento + coche + comidas + atracciones)",
        "Alcohol cero al volante — límite legal: 0,0 g/L. Ninguna copa antes de conducir",
        "Carreteras de montaña con baches y animales sueltos: conducir siempre con calma",
        "Vestimenta discreta en monasterios: hombros cubiertos, faldas/pantalones bajo la rodilla",
      ]),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // POR QUÉ RUMANÍA
      // ═══════════════════════════════════════════════════════
      sectionHead("Por qué Rumanía"),
      body("Rumanía es uno de los destinos más fascinantes y menos masificados de Europa. Un país donde el pasado medieval convive con una naturaleza salvaje y desbordante, donde los pueblos conservan tradiciones centenarias y la gastronomía no ha sido domesticada para el turismo. Un viaje en coche por Rumanía es, ante todo, una aventura de descubrimiento."),
      empty(80),
      body("A diferencia de sus vecinos centroeuropeos, Rumanía todavía sorprende. Las carreteras serpentean entre bosques de hayas, castillos sobre acantilados y monasterios con frescos del siglo XV pintados en sus paredes exteriores. Los precios son extraordinariamente bajos para los estándares europeos occidentales, y la hospitalidad rumana —especialmente en zonas rurales— es genuina y cálida."),
      empty(120),
      highlight("Rumanía cuenta con más de 500 iglesias y castillos históricos, 13 sitios del Patrimonio Mundial de la UNESCO y los bosques vírgenes más extensos de Europa fuera de Rusia."),
      empty(160),
      infoBox("5 razones para elegir Rumanía como próximo destino", [
        "Naturaleza extraordinaria: los Cárpatos, el Delta del Danubio y bosques primigenios",
        "Historia medieval viva: Sighișoara, Sibiu, Brașov, Sinaia y docenas de castillos",
        "Gastronomía auténtica y asequible: cocina de herencia latina, eslava y balcánica",
        "Sin masificación: fuera de los grandes circuitos turísticos europeos",
        "Road trip perfecto: carreteras espectaculares como la Transfăgărășan o la Transalpina",
      ]),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // EL CLIMA EN JULIO
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("El Clima en Julio"),
      body("Julio es un mes de contrastes en Rumanía. Mientras Bucarest puede registrar temperaturas de hasta 35 °C, la cima de la Transfăgărășan a 2.042 metros puede bajar de los 10 °C con viento y niebla el mismo día. Entender esta diferencia es clave para equiparse bien y planificar el ritmo del viaje."),
      empty(120),
      climateTable(),
      empty(160),

      subsectionHead("Qué esperar en la ruta"),
      body("Las ciudades de Transilvania (Brașov, Sibiu, Sighișoara) son las más agradables en julio: días soleados de 25-28 °C sin el agobio de Bucarest. El norte —Maramureș y Bucovina— es algo más fresco y húmedo, con lluvias vespertinas frecuentes que dejan el paisaje de un verde intenso. Bucovina es la región con más precipitaciones del país, especialmente en la segunda quincena de julio."),
      empty(80),
      body("La alta montaña merece atención especial. En el Lago Bâlea (Transfăgărășan, 2.034 m) el termómetro rara vez supera los 15 °C incluso en agosto, y la niebla puede cerrar la visibilidad en cuestión de minutos. Las tormentas de tarde son habituales: lo ideal es cruzar el punto más alto antes de las 14h."),
      empty(120),
      infoBox("Consejos de equipaje derivados del clima", [
        "Capas, no ropa de verano exclusivamente: la diferencia entre Bucarest y la Transfăgărășan puede ser de 20 °C el mismo día",
        "Chubasquero o impermeable ligero: lluvia vespertina frecuente en el norte (Maramureș, Bucovina)",
        "Protector solar factor 50: a 2.000 metros la radiación UV es significativamente mayor que en las ciudades",
        "Calzado impermeable para senderismo: los senderos de las gargantas de Bicaz y los alrededores del Lago Bâlea pueden estar húmedos",
      ]),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // DURACIÓN RECOMENDADA
      // ═══════════════════════════════════════════════════════
      sectionHead("Duración Recomendada"),
      body("Para un viajero en solitario que quiere combinar naturaleza, historia, gastronomía y vida rural, la duración ideal de un road trip por Rumanía son 12 días completos en el país (13 días con vuelos desde Alicante). Esto permite:"),
      empty(80),
      bullet("Recorrer las seis regiones del viaje: Muntenia, Oltenia, Transilvania, Maramureș, Moldova y Bucovina"),
      bullet("Conducir la Transfăgărășan con tiempo para paradas y senderismo"),
      bullet("Explorar en profundidad dos o tres ciudades medievales: Brașov, Sibiu, Sighișoara"),
      bullet("Sumergirse en la vida rural del norte del país sin prisas"),
      bullet("Tener 1-2 días de margen para imprevistos, desvíos espontáneos o descanso"),
      empty(120),
      highlight("Con 7-10 días se puede hacer un recorrido de lujo por Transilvania, pero sería una pena no llegar a Maramureș y Bucovina, que son la auténtica joya del país."),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // LAS SEIS REGIONES
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Las Seis Regiones del Viaje"),
      body("El itinerario cruza seis regiones históricas de Rumanía, cada una con su propio carácter, arquitectura y tradición gastronómica. Conocerlas antes de salir ayuda a entender por qué cada etapa del viaje se siente tan diferente de la anterior."),
      empty(120),

      subsectionHead("Muntenia — La gran llanura y la capital"),
      body("Muntenia (la Valaquia oriental) es la región más poblada de Rumanía y su corazón político. Bucarest, su capital, combina grandiosidad comunista —el monumental Palacio del Parlamento de Ceaușescu— con barrios art nouveau del siglo XIX y una escena gastronómica y cultural sorprendentemente vibrante. Al norte, Curtea de Argeș conserva el monasterio episcopal más bello del país. El viaje empieza y termina en Muntenia."),
      empty(120),

      subsectionHead("Oltenia — El paso del Danubio al corazón del país"),
      body("Oltenia, la Valaquia occidental, es la región de tránsito natural entre el sur del país y los Cárpatos. El valle del río Olt —que la atraviesa de norte a sur— es uno de los corredores históricos más importantes de Rumanía. En sus márgenes se levanta el Monasterio de Cozia (siglo XIV), uno de los conjuntos medievales mejor conservados del sur del país. La cerámica de Horezu, patrimonio inmaterial de la UNESCO, es la artesanía más representativa de la región."),
      empty(120),

      subsectionHead("Transilvania — La tierra de los castillos"),
      body("El corazón histórico de Rumanía es también su región más conocida internacionalmente. Durante siglos fue una región autónoma con una compleja mezcla de poblaciones valaca, húngara y alemana sajona, lo que explica la variedad de influencias que se aprecian en sus ciudades. La Cordillera de los Cárpatos la rodea en forma de arco, aislándola del resto del país y dándole un carácter propio inconfundible. Ciudades clave del itinerario: Brașov, Sibiu, Sighișoara, Cluj-Napoca."),
      empty(120),

      subsectionHead("Maramureș — El museo vivo de la tradición"),
      body("Esta región del extremo norte, en la frontera con Ucrania, es el único lugar de Europa donde la vida rural medieval se ha conservado de manera orgánica, sin artificios. Los vecinos siguen arando con bueyes, construyendo con madera y celebrando las fiestas patronales con trajes bordados a mano. Sus iglesias de madera —ocho de ellas declaradas Patrimonio de la Humanidad— son la expresión más pura de una arquitectura que no necesitó piedra para alcanzar la trascendencia. Maramureș no es un parque temático: es la vida real de un territorio que el tiempo ha tratado con suavidad."),
      empty(120),

      subsectionHead("Moldova — Los Cárpatos orientales y las gargantas"),
      body("La Moldova rumana (distinta del país independiente del mismo nombre) es la región que une el norte con el este del país. Sus Cárpatos orientales ofrecen algunos de los paisajes más dramáticos del viaje: las Gargantas de Bicaz —paredes calizas de 300 metros sobre la carretera— y el misterioso Lago Rojo, formado por un corrimiento de tierras en 1837. Es una región menos frecuentada por el turismo internacional, lo que garantiza autenticidad y precios especialmente bajos."),
      empty(120),

      subsectionHead("Bucovina — El museo del fresco al aire libre"),
      body("La Bucovina histórica (hoy dividida entre Rumanía y Ucrania) fue durante el siglo XV el centro del arte religioso moldavo. Los príncipes mandaron pintar el exterior de sus monasterios como 'biblia de los pobres' para los fieles analfabetos. El resultado —murales de miles de metros cuadrados a la intemperie que han sobrevivido más de 500 años— es una de las maravillas artísticas menos conocidas del mundo. El azul de Voroneț, imposible de reproducir con los pigmentos actuales, es el símbolo visual de toda la región."),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // ITINERARIO
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Itinerario: 12 Días en Coche por Rumanía"),
      body("El siguiente itinerario está diseñado para un viajero en solitario que sale desde Alicante, recoge un coche de alquiler en el Aeropuerto de Bucarest (OTP) el segundo día y realiza un circuito de ida y vuelta. Las distancias diarias están pensadas para que el viaje sea cómodo: entre 100 y 200 km por día como máximo en los tramos exigentes. La última noche se duerme junto al aeropuerto OTP para el vuelo de regreso a Alicante a las 06:45."),
      empty(120),

      // ── DÍA 1 ──
      dayBlock(1, "El Caos que Enamora", "Bucarest — Primera noche en la capital", "vuelo · llegada 14:40", "Muntenia",
        "Bucarest: de fortaleza a 'París de los Balcanes'",
        "Fundada como fortaleza en el siglo XV y nombrada según la leyenda por el pastor Bucur, Bucarest se convirtió en capital de Valaquia en 1659. Durante el siglo XIX vivió su época dorada: los grandes bulevares y los palacetes del arte nouveau otorgaron a la ciudad el apodo de 'París de los Balcanes'. El régimen de Ceaușescu demolió buena parte del casco histórico en los años 80 para construir el monumental Palacio del Parlamento. La Revolución de diciembre de 1989 puso fin a la dictadura."
      ),
      empty(100),
      body("Vuelo Alicante (ALC) → Bucarest (OTP) con Wizz Air. Salida a las 10:15, llegada a las 14:40 (3h 25 min). Salir del aeropuerto lleva unos 30-40 minutos entre equipaje y controles; estarás en el hotel hacia las 15:30-16:00. Tomar un Uber o taxi oficial hasta el alojamiento — mañana se recoge el coche. No conducir por el centro de Bucarest si se puede evitar: el tráfico es un mundo aparte."),
      empty(80),
      body("Primera parada: el Palacio del Parlamento. El exterior es gratuito; el interior solo se visita con visita guiada, que conviene reservar con antelación (~45 lei). Es el segundo edificio más grande del mundo: Nicolae Ceaușescu lo mandó construir en 1984 y costó la demolición de un barrio histórico entero. Imprescindible verlo aunque sea desde fuera."),
      empty(80),
      body("Desde el Parlamento, bajar por el Bulevardul Unirii: los 3,5 km de bulevar que unen el Palacio con la Piața Alba Iulia, diseñados por Ceaușescu como respuesta al Champs-Élysées. Un paseo de 20-30 minutos que lleva directamente al centro histórico."),
      empty(80),
      body("Tarde y noche: primer contacto con Bucarest. El pasaje de cristal Macca-Villacrosse —de 1891, con su techo de vidrio amarillo— es uno de los rincones más fotogénicos de la ciudad. El barrio de Lipscani (el casco histórico) es el lugar ideal para una primera cerveza rumana (Ursus o Timișoreana) y una cena tranquila."),
      empty(80),
      bullet("Alojarse en el barrio de Dorobanți o Floreasca: tranquilos, con buena conexión al aeropuerto y aparcamiento para el día siguiente."),
      bullet("Cena recomendada: Caru' cu Bere, la cervecería histórica más bella de Bucarest (1879), con interior neogótico espectacular."),
      empty(80),
      momentDelDia("Cruzar el pasaje Macca-Villacrosse a última hora de la tarde, con la luz de julio filtrándose a través del techo de vidrio amarillo. Bucarest huele a tilo en julio. Ese olor lo recuerdas semanas después."),
      empty(160),

      // ── DÍA 2 ──
      dayPageBreak(),
      dayBlock(2, "La Puerta de los Cárpatos", "Bucarest → Sinaia → Brașov", "154 km · 2h 30 min", "Muntenia · Transilvania",
        "Sinaia: el retiro real que nació de un monasterio",
        "El Valle de Prahova fue durante siglos la puerta natural entre Valaquia y Transilvania. El Monasterio de Sinaia fue fundado en 1695 por el noble Mihai Cantacuzino, quien lo dedicó al Monte Sinaí donde había peregrinado. Cuando el rey Carol I descubrió el valle en 1866, mandó construir el Castillo de Peleș, convirtiendo el lugar en el 'Biarritz rumano'. Brașov, fundada en 1211 por los Caballeros Teutónicos, fue durante siglos el mayor centro comercial del sudeste de Europa."
      ),
      empty(100),
      body("Mañana: recoger el coche de alquiler en el Aeropuerto OTP (reservar previamente con Europcar, Avis o Budget). Primera conducción por Rumanía: carretera DN1 hacia el norte, subiendo suavemente hacia los Cárpatos por el Valle de Prahova."),
      empty(80),
      body("Parada imprescindible en Sinaia: el Castillo de Peleș (1883), residencia de verano del rey Carol I. Un palacio de estilo neo-renacimiento alemán completamente fuera de lugar en los Cárpatos —y precisamente por eso, irresistible. Entradas en efectivo: 40-80 lei según el recorrido."),
      empty(80),
      body("Por la tarde, llegar a Brașov. La ciudad más visitada de Transilvania ofrece una arquitectura saxona impecable, restaurantes excelentes y una ubicación perfecta como base para los próximos días."),
      empty(80),
      bullet("El Castillo de Bran ('Drácula') está a 28 km de Brașov. Conviene tener claras las expectativas: su vínculo real con Vlad Țepeș es mínimo y comercialmente exagerado, la visita es cara y masificada, y el castillo en sí es pequeño. Peleș (que verás mañana) lo supera en todos los sentidos. Bran merece una parada rápida solo si la curiosidad puede más —y sabiendo lo que es."),
      bullet("Activar la rovinieta antes de salir del aeropuerto: www.roviniete.ro o en cualquier gasolinera."),
      bullet("Cena recomendada: Restaurante Sergiana (cocina tradicional rumana, especialidad sarmale)."),
      empty(80),
      momentDelDia("El primer momento en que el Castillo de Peleș aparece entre los abetos: ves un palacio de cuento en mitad de los Cárpatos y no entiendes cómo nadie te había avisado de que algo así existía en Rumanía."),
      empty(160),

      // ── DÍA 3 ──
      dayPageBreak(),
      dayBlock(3, "Entre Osos y Campanarios", "Brașov — La ciudad sajona", "día de base", "Transilvania",
        "Brașov: Kronstadt, la ciudad que tres pueblos construyeron",
        "Los sajones llamados por el rey húngaro Géza II en el siglo XII levantaron Kronstadt como bastión comercial y militar. La Iglesia Negra —renombrada así tras el gran incendio de 1689— es la mayor iglesia gótica de Europa central al sur de los Alpes. Junto a sajones y húngaros, los valacos habitaban el barrio de Șchei, fuera de las murallas. Tras la caída del comunismo, la mayoría de los sajones transilvanos emigró a Alemania entre 1990 y 2000."
      ),
      empty(100),
      body("Día completo en Brașov. La agenda de mañana incluye la Iglesia Negra del siglo XIV (la mayor iglesia gótica del sur de Europa central), la Plaza del Consejo (Piața Sfatului) y el cartel de BRAȘOV en la colina. La calle Sforii es considerada una de las calles más estrechas de Europa (1,32 m en su punto más angosto)."),
      empty(80),
      body("Por la tarde, experiencia estrella: excursión de avistamiento de osos en los Cárpatos. Varios operadores locales organizan salidas de 2-3 horas a los escondites en el bosque donde los osos pardos bajan a alimentarse al atardecer. Rumanía alberga la mayor población de osos pardos de Europa continental, con más de 6.000 ejemplares. Coste: 45-60 € por persona."),
      empty(80),
      bullet("La Ciudadela de Râșnov (a 16 km) merece una visita rápida de 1 hora: fortaleza campesina del siglo XIII con vista panorámica excepcional."),
      bullet("Alojarse 2 noches en Brașov para aprovechar mejor los alrededores."),
      empty(80),
      infoBox("Opcional: Muntele Omu (2.505 m)", [
        "El pico más alto de los Bucegi, accesible desde Bușteni (45 km de Brașov) en combinación de coche y telecabina",
        "Telecabina Bușteni → Babele (1.680 m) + senderismo hasta la cumbre: 1,5-2 h de subida",
        "Arriba hay una estación meteorológica con refugio y vistas panorámicas sobre toda la cadena de los Cárpatos",
        "Recomendación: ir temprano (antes de las 9h) para evitar colas en la telecabina en julio",
        "Imprescindible: ropa de abrigo y chubasquero aunque en Brașov haga calor — en la cima puede bajar a 10 °C",
      ]),
      empty(80),
      momentDelDia("Sentarte al atardecer en Piața Sfatului con una Ursus fría mientras las campanas de la Iglesia Negra empiezan a sonar. Brașov es de las ciudades que te devuelven las ganas de viajar."),
      empty(160),

      // ── DÍA 4 ──
      dayPageBreak(),
      dayBlock(4, "La Ruta que Conquistó el Mundo", "La Carretera Transfăgărășan", "133 km · 3h 30 min", "Transilvania",
        "La Transfăgărășan: la carretera que nació del miedo",
        "Ceaușescu ordenó construir la Transfăgărășan entre 1970 y 1974 tras la invasión soviética de Checoslovaquia en 1968: necesitaba una ruta militar alternativa a través de los Cárpatos. La construcción requirió 6.000 toneladas de dinamita y costó la vida a decenas de trabajadores. Jeremy Clarkson la declaró 'la mejor carretera del mundo' en Top Gear (2009). Solo está abierta de junio a octubre."
      ),
      empty(100),
      body("Este es el día más esperado del viaje. La Transfăgărășan (DN7C) asciende hasta los 2.042 metros de altitud, pasando por la Cascada Bâlea, el impresionante Lago Bâlea entre nieves perpetuas, y el Lago Vidraru con su enorme presa. Un espectáculo de ingeniería y naturaleza a partes iguales."),
      empty(80),
      bullet("Salir temprano (antes de las 8h) para evitar la cola de coches en el tramo norte."),
      bullet("Llevar abrigo: en la cima puede hacer frío incluso en julio."),
      bullet("No hay gasolineras en la ruta: repostar antes de entrar."),
      bullet("Opción de alojamiento en la Cabaña Bâlea Lac, a 2.034 m de altitud."),
      empty(80),
      momentDelDia("Aparcar en el collado de Bâlea Lac (2.034 m) y apagar el motor. El silencio de alta montaña, el olor a hierba mojada y el lago glaciar a tus pies. Ninguna foto lo captura. Esto es para lo que viniste."),
      empty(120),
      motoBox("Si quieres vivir la Transfăgărășan de verdad, hazla en moto. La empresa Transylvania Moto Experience (Brașov) organiza un tour guiado de un día completo con salida y recogida en el hotel, moto y guía incluidos. La carretera fue diseñada para el movimiento rápido de vehículos militares, y en dos ruedas cada curva y cada recta cobran una dimensión completamente diferente. Precio aproximado: 150-180 €. Reserva previa imprescindible en su web. Esta es la alternativa al día en coche, no una actividad adicional."),
      empty(160),

      // ── DÍA 5 ──
      dayPageBreak(),
      dayBlock(5, "El Otro Lado", "Transfăgărășan → Curtea de Argeș → Sibiu", "100 km · 2h 15 min", "Transilvania · Muntenia · Oltenia",
        "Sibiu: la capital de la Sajonia transilvana",
        "Sibiu (Hermannstadt en alemán) fue fundada en el siglo XII por colonos sajones convocados por la corona húngara. Entre 1692 y 1791 fue capital del principado de Transilvania. Su detalle más singular son los 'ojos de Sibiu': ventanas de buhardilla con forma de párpado inclinado que permitían ventilar los graneros y vigilar la calle. En 2007 fue Capital Europea de la Cultura junto con Luxemburgo."
      ),
      empty(100),
      body("Mañana con posibilidad de senderismo en los alrededores de Bâlea Lac antes del descenso. El tramo sur de la Transfăgărășan baja hacia Curtea de Argeș, donde vale la pena parar en el Monasterio Episcopal (siglo XVI), protagonista de la famosa leyenda del Maestru Manole."),
      empty(80),
      body("Después, dirección Sibiu: la ciudad más sorprendente de Rumanía. Conocida como 'la ciudad de los ojos' por las ventanas de las buhardillas que parecen mirarte fijamente, Sibiu fue Capital Europea de la Cultura en 2007 y su casco histórico es uno de los mejor conservados de la región."),
      empty(80),
      bullet("Aparcar en los aparcamientos perimetrales de Sibiu y explorar a pie el centro histórico."),
      bullet("Cena en el restaurante Hermania: excelente fusión de cocina alemana-sajona y rumana."),
      empty(80),
      infoBox("Opcional: Monasterio de Cozia", [
        "En la bajada de la Transfăgărășan hacia el sur, el valle del río Olt guarda uno de los monumentos medievales mejor conservados de Rumanía: el Monasterio de Cozia, fundado en 1386 por el príncipe Mircea el Viejo.",
        "La iglesia principal conserva frescos del siglo XIV prácticamente intactos y el cuerpo del propio Mircea reposa en su interior — todo un símbolo de la resistencia de Oltenia frente al avance otomano.",
        "La visita añade unos 15 minutos de desvío. Está señalizada desde la carretera DN7 y merece la parada aunque sea breve.",
      ]),
      empty(80),
      momentDelDia("Llegar a Sibiu al atardecer y encontrar la Piața Mare iluminada. Las ventanas de las buhardillas —los 'ojos de Sibiu'— te observan desde todos los tejados. Tienes la sensación de que la ciudad te está evaluando."),
      empty(160),

      // ── DÍA 6 ──
      dayPageBreak(),
      dayBlock(6, "La Ciudad de los Ojos", "Sibiu — Profundidad y calma", "día de base", "Transilvania",
        "El legado sajón: por qué Sibiu tiene tres plazas y 'ojos' en los tejados",
        "La topografía de Sibiu —ciudad alta (intramuros) y ciudad baja (comercial)— refleja la estructura social medieval. Las ventanas de buhardilla con forma de ojo son un elemento funcional: un ojo de buey inclinado hacia abajo para ventilar y vigilar. Tras la reunificación alemana, la mayor parte de la comunidad sajona (unos 40.000 habitantes en 1989) emigró a Alemania en apenas diez años."
      ),
      empty(100),
      body("Día completo para explorar Sibiu sin prisas. Los tres niveles de plazas conectadas (Piața Mare, Piața Mică y Piața Huet) forman uno de los conjuntos urbanos más elegantes de Transilvania. El Puente de los Mentirosos (Podul Minciunilor), el primero de hierro fundido de Rumanía (1859), conecta la plaza mayor con el barrio inferior."),
      empty(80),
      body("El Complejo ASTRA del Museo Nacional, a las afueras de la ciudad, es un museo al aire libre de 96 hectáreas con más de 300 construcciones tradicionales trasladadas y reconstruidas. Es uno de los mejores museos etnográficos de Europa y merece una mañana entera — y también uno de los mejores lugares para llevarse artesanía auténtica a casa."),
      empty(80),
      bullet("El Museo Brukenthal (palacio barroco del siglo XVIII) alberga una colección de pintura europea sorprendentemente buena."),
      bullet("El mercado cubierto junto a Piața Mare: queso de burduf, embutidos ahumados y pálinca artesanal."),
      empty(80),
      momentDelDia("Una mañana entera vagando por el Museo ASTRA sin prisa, entre casas de madera del siglo XVIII trasladadas piedra a piedra. La sensación de haber encontrado un Rumanía que el mundo no sabe que existe."),
      empty(160),

      // ── DÍA 7 ──
      dayPageBreak(),
      dayBlock(7, "Donde Nació el Mito", "Sibiu → Biertan → Sighișoara", "130 km · 2h 30 min", "Transilvania",
        "Sighișoara: la cuna de Vlad III y la ciudadela que el tiempo no tocó",
        "Fundada hacia 1150 por colonos sajones, Sighișoara alcanzó su esplendor como ciudad comercial en los siglos XIV y XV. En 1431 nació aquí Vlad III Drăculea, príncipe de Valaquia cuya ferocidad contra los invasores otomanos —empalaba a sus enemigos— inspiró la novela de Bram Stoker. La Torre del Reloj data del siglo XIV. Lo excepcional de Sighișoara es que su ciudadela amurallada sigue siendo residencial: vecinos reales viven hoy en las mismas casas medievales. La UNESCO la declaró Patrimonio de la Humanidad en 1999."
      ),
      empty(100),
      body("Antes de llegar a Sighișoara, la ruta pasa junto a Biertan: un pueblo sajón de 2.500 habitantes con una de las iglesias fortificadas más grandes y mejor conservadas de Transilvania (Patrimonio UNESCO desde 1993). El desvío desde la DN14 es mínimo y la parada merece una hora: la iglesia-fortaleza del siglo XV domina el valle desde lo alto y el pueblo a sus pies está casi intacto, con muy poco turismo comparado con Sighișoara o Brașov."),
      empty(80),
      body("Sighișoara es probablemente la ciudad medieval mejor conservada del mundo. Su ciudadela superior —intramuros— sigue habitada, algo excepcional en Europa. Aquí nació en 1431 Vlad III, el histórico príncipe valaco que inspiró el mito del Conde Drácula."),
      empty(80),
      body("La Torre del Reloj (siglo XIV), la Iglesia de la Colina y las callejuelas empedradas con casas de colores constituyen uno de esos escenarios que parecen pintados. No perderse el Pasaje de los Estudiantes, cubierto y con escaleras de madera, que sube a la ciudadela."),
      empty(80),
      bullet("Biertan: aparcar abajo en el pueblo y subir a pie (10 minutos). Entrada a la iglesia: ~15 lei. Mañana, antes de que llegue el calor."),
      bullet("Sighișoara está menos masificada que Brașov o Sibiu. Es una parada de medio día perfecta."),
      bullet("Si coincide con el Festival Medieval de Sighișoara (segunda semana de julio), una de las mejores experiencias del viaje."),
      bullet("Opcional (día maratoniano): Castillo de Corvin en Hunedoara. Está ~130 km al oeste de Sibiu —en dirección contraria al resto del viaje—, lo que supone unos 250 km extra y 4-5 horas de volante adicionales. Es uno de los castillos góticos más impresionantes de Europa, pero requiere planificarlo como día específico o como parada entre Sibiu y Cluj con pernoctación en ruta.", false),
      empty(80),
      momentDelDia("Subir el Pasaje de los Estudiantes en Sighișoara al anochecer, cuando los turistas ya se han ido y las callejuelas de la ciudadela están casi vacías. La luz de los faroles sobre el adoquín y el silencio. Sighișoara tiene otra dimensión cuando te quedas solo en ella."),
      empty(160),

      // ── DÍA 8 ──
      dayPageBreak(),
      dayBlock(8, "El Corazón Joven de Transilvania", "Sighișoara → Salina Turda → Cluj-Napoca", "120 km · 2h 15 min", "Transilvania",
        "Cluj-Napoca: cuatro nombres y veinte siglos de historia",
        "La ciudad fue fundada como colonia romana con el nombre de Napoca en el siglo II d.C. Los húngaros la llamaron Kolozsvár, los alemanes Klausenburg y los rumanos Cluj. Esta multiplicidad de nombres refleja siglos de convivencia y tensión entre comunidades. Capital histórica del principado de Transilvania, hoy es la segunda ciudad de Rumanía con más de 320.000 habitantes y más de 100.000 estudiantes universitarios."
      ),
      empty(100),
      body("Cluj-Napoca es la ciudad más vibrante de Rumanía. Capital intelectual del país, sede del mayor festival de música de Europa del Este (UNTOLD, en agosto) y ciudad con una energía juvenil que contrasta con las ciudades medievales de días anteriores. Su casco histórico, algo más irregular que el de Sibiu o Brașov, tiene su propio encanto auténtico."),
      empty(80),
      body("La Plaza Unirii, con la majestuosa Catedral de San Miguel (siglo XIV, gótico tardío) y la estatua ecuestre de Matías Corvino, es el centro neurálgico de la ciudad. El Jardín Botánico de Cluj (uno de los mayores de Europa del Este) es ideal para un paseo tranquilo antes del calor del mediodía."),
      empty(80),
      bullet("Salina Turda — La mina de sal del siglo XI reconvertida en parque subterráneo espectacular. Está directamente en la ruta Sighișoara→Cluj, a 30 km antes de llegar a la ciudad. Entrada: ~60 lei. Llegar antes de las 10h o reservar entrada online para evitar colas en julio.", true),
      bullet("Desvío opcional desde Cluj: Cheile Turzii (Gargantas de Turda, 30 km al sur). Desfiladero calizo con paredes de hasta 300 m. Caminata de 1,5-2 horas sobre puentes colgantes."),
      bullet("Cena en el barrio universitario: restaurantes locales de excelente calidad-precio."),
      empty(80),
      momentDelDia("Bajar a la Salina Turda y encontrarte bajo tierra con un lago subterráneo, una noria del siglo XIX y el silencio absoluto a 120 metros de profundidad. El mundo de arriba desaparece completamente."),
      empty(160),

      // ── DÍA 9 ──
      dayPageBreak(),
      dayBlock(9, "El Norte que el Tiempo Olvidó", "Cluj-Napoca → Maramureș", "200 km · 3h 30 min", "Transilvania · Maramureș",
        "Maramureș: la región que sobrevivió entre imperios",
        "Esta región fronteriza del norte fue durante siglos disputada entre el reino húngaro, Moldavia y Valaquia. Sus habitantes conservaron una notable autonomía frente a los grandes imperios, lo que explica que sus tradiciones, arquitectura en madera y ritos se hayan preservado de forma tan orgánica: no como un parque temático, sino como forma real de vida."
      ),
      empty(100),
      body("Hoy el paisaje cambia radicalmente. Dejar atrás las ciudades medievales de Transilvania para adentrarse hacia el norte es uno de los viajes en el tiempo más sorprendentes de Europa. A medida que la carretera sube hacia Maramureș, los pueblos se vuelven más pequeños, los carros tirados por caballos más frecuentes y la modernidad más escasa."),
      empty(80),
      body("Las carreteras en el norte de Rumanía pierden calidad: muchas son de dos carriles sin arcén y con baches. Conducir con calma y sin prisa. El propio trayecto es ya parte de la experiencia."),
      empty(80),
      bullet("Alojarse en una pensiune local. La hospitalidad en Maramureș es legendaria."),
      bullet("Cena incluida en la mayoría de pensiuneas: comida casera de kilómetro cero, un lujo."),
      bullet("El mercado semanal de Sighetul Marmației (sábados) merece ajustar el itinerario si se puede."),
      empty(80),
      momentDelDia("Detenerte en un pueblo de Maramureș al atardecer y ver a un vecino llegar a casa en un carro tirado por caballos. No hay ironía, no hay turismo. Eso es simplemente cómo llega a casa."),
      empty(160),

      // ── DÍA 10 ──
      dayPageBreak(),
      dayBlock(10, "Madera, Madera y Madera", "Maramureș — Iglesias centenarias y tradición viva", "120 km · 3h de ruta", "Maramureș",
        "Las iglesias de madera: la respuesta a una prohibición",
        "Las iglesias de madera de Maramureș son, en parte, consecuencia de una imposición del Imperio Austro-Húngaro que prohibió a los ortodoxos construir iglesias en piedra. La comunidad respondió con torres de madera de roble que en algunos casos alcanzan los 54 metros de altura, construidas sin un solo clavo. Ocho de ellas están declaradas Patrimonio de la Humanidad por la UNESCO desde 1999."
      ),
      empty(100),
      body("Maramureș es la región más auténtica de Rumanía. Un valle encajado en los Cárpatos septentrionales donde las tradiciones medievales no son una recreación turística sino parte del día a día. Las iglesias de madera son la máxima expresión de una arquitectura vernácula excepcional."),
      empty(80),
      body("El Monasterio de Bârsana y la Iglesia de Desești tienen frescos interiores de una calidad pictórica asombrosa. El pueblo de Breb es considerado el más auténtico de toda Rumanía: casas de madera con portones tallados, huertos, y vecinos que todavía van a los mercados en traje tradicional."),
      empty(80),
      body("La joya excéntrica de la región es el Cementerio Alegre (Cimitirul Vesel) de Săpânța: un camposanto donde cada lápida tiene una escultura y un poema en primera persona del difunto, a menudo con humor negro y colores vivos. Una celebración de la vida iniciada por el tallista Stan Ioan Pătraș en 1935."),
      empty(80),
      bullet("La Vía Estrecha del Valle de Vaser (Mocănița): tren a vapor de principios del siglo XX que sube por el valle durante 4 horas. Solo en temporada turística — RESERVAR CON SEMANAS DE ANTELACIÓN desde España. En julio las plazas se agotan. Web: cffviseudesus.ro", true),
      bullet("El mercado semanal de Sighetul Marmației (los sábados): productos locales, artesanía y vida real."),
      empty(80),
      momentDelDia("Ver el Cementerio Alegre de Săpânța y leer los poemas en las lápidas. Los rumanos llevan siglos riéndose de la muerte. Ese detalle lo dice todo sobre este país."),
      empty(160),

      // ── DÍA 11 ──
      dayPageBreak(),
      dayBlock(11, "El Desfiladero que Corta la Respiración", "Cheile Bicazului y el Lago Rojo", "180 km · 3h 45 min", "Maramureș · Moldova · Bucovina",
        "El Lago Rojo: cuando una montaña decidió taponar un río",
        "En agosto de 1837, un desprendimiento del monte Ghilcoș bloqueó el río Bicaz, formando el Lago Rojo. El nombre proviene del color rojizo que adquirían sus aguas teñidas por el óxido de hierro de los afluentes. Los troncos del bosque inundado aún asoman sobre la superficie, creando una imagen fantasmal. Las Gargantas de Bicaz se formaron a lo largo de millones de años por erosión del río sobre la roca caliza, con paredes de hasta 300 metros."
      ),
      empty(100),
      body("El trayecto de Maramureș al sur atraviesa los Cárpatos por carreteras de montaña de gran belleza. La gran parada del día son las Gargantas de Bicaz (Cheile Bicazului): uno de los desfiladeros más impresionantes de Rumanía, con paredes calizas verticales de hasta 300 metros que se cierran sobre la carretera creando un pasillo de roca casi claustrofóbico durante 5 km."),
      empty(80),
      body("A pocos kilómetros, el Lago Rojo (Lacul Roșu): formado en 1837 por un desprendimiento, con los troncos muertos de la antigua arboleda asomando fantasmales por la superficie. Un paisaje completamente diferente a todo lo visto hasta ahora."),
      empty(80),
      bullet("Hacer la ruta a pie por el interior de las gargantas: el sendero junto al río es accesible y permite ver las paredes desde abajo."),
      bullet("Los puestos locales a la entrada venden artesanía, quesos y fruta: un descanso perfecto a mitad de jornada."),
      bullet("Al caer la tarde, continuar hacia el norte (unos 120 km) para alojarse ya en la zona de Gura Humorului o Suceava, en el corazón de Bucovina. Así el día siguiente empieza sin prisas de carretera."),
      empty(80),
      momentDelDia("Estar dentro de las Gargantas de Bicaz a primera hora, antes de que lleguen los autobuses, con el río a los pies y 300 metros de roca vertical cerrándose sobre tu cabeza. Es el paisaje más claustrofóbico y más bello del viaje."),
      empty(160),

      // ── DÍA 12 ──
      dayPageBreak(),
      dayBlock(12, "El Azul que No Se Puede Fotografiar", "Bucovina — Voroneț y Sucevița — Regreso a Bucarest", "430 km · 6h de conducción", "Bucovina · Muntenia",
        "Los monasterios pintados: una biblia al aire libre",
        "Los monasterios pintados de Bucovina fueron construidos entre 1487 y 1583 por los príncipes moldavos, especialmente Ștefan cel Mare (Esteban el Grande), quien mandó levantar una iglesia tras cada victoria militar sobre los otomanos. Los frescos exteriores eran la 'biblia de los pobres' para fieles analfabetos. El azul de Voroneț, obtenido con lapislázuli y otros minerales, sigue siendo imposible de reproducir con exactitud."
      ),
      empty(100),
      body("El último día completo: Bucovina. Madrugar para llegar a los monasterios pintados antes de los grupos de turistas. El Monasterio de Voroneț (1487) —llamado la 'Capilla Sixtina de Oriente'— tiene el azul más intenso que existe en la pintura al fresco. La representación del Juicio Final en su fachada ocupa 6 metros de altura y es sencillamente abrumadora."),
      empty(80),
      body("El Monasterio de Sucevița (finales del XVI) es el más grande y mejor preservado del conjunto. Sus frescos asombran por su viveza tras siglos a la intemperie, especialmente la Escalera del Paraíso de la fachada norte. Gracias a haber dormido en Bucovina la noche anterior, los monasterios pueden visitarse sin prisa y sin madrugar en exceso."),
      empty(80),
      body("Después de Sucevița, iniciar el regreso a Bucarest por la DN2/E85: unos 430 km y 5-6 horas de conducción. Llegar al aeropuerto OTP al caer la noche, devolver el coche de alquiler en las oficinas del terminal y hacer el check-in en un hotel junto al aeropuerto. Es la última noche del viaje."),
      empty(80),
      bullet("Los monasterios son lugares de culto activos: vestimenta discreta obligatoria. Entrada: ~10 lei por monasterio (~2 €)."),
      bullet("Contratar una guía local para los frescos: sin explicación, se pierden detalles fascinantes."),
      bullet("Salir de Bucovina antes de las 14h para llegar a OTP con margen y sin el estrés del tráfico de accesos a Bucarest."),
      bullet("Hotel recomendado: zona Aeropuerto OTP (Otopeni). Opciones económicas a 5-10 minutos del terminal."),
      bullet("Vuelo de regreso: Wizz Air OTP → ALC, salida 06:45, llegada Alicante 09:30. Estar en el aeropuerto a las 04:45. Facturar online desde 48h antes — recomendable hacerlo la tarde anterior desde el hotel."),
      empty(80),
      momentDelDia("Plantarte delante del Monasterio de Voroneț a primera hora de la mañana, sin nadie más. El azul de los frescos exteriores con la luz baja del amanecer es una de esas imágenes que no se parecen a ninguna fotografía que hayas visto antes. El color existe. Está ahí. Y luego vuelves a casa."),
      // ═══════════════════════════════════════════════════════
      // GASTRONOMÍA
      // ═══════════════════════════════════════════════════════
      sectionHead("Gastronomía Rumana", true),
      body("La cocina rumana es hija de su historia: un territorio que fue frontera entre el Imperio Romano, el otomano y el austro-húngaro durante siglos, y que recibió influencias eslavas, latinas, turcas y magiares que se fueron sedimentando en una gastronomía propia, contundente y de una generosidad que sorprende al viajero occidental. No hay cocina de fusión aquí: hay cocina de supervivencia que se volvió deliciosa."),
      empty(80),
      body("La base es la tierra: cerdo criado en el campo, cordero de las laderas de los Cárpatos, maíz para la mămăligă, coles fermentadas para los sarmale, setas del bosque, ciruelas para la pálinca. Todo tiene un origen claro y una función en la dieta de una población que vivió durante siglos de lo que producía. Esa coherencia se siente en el plato."),
      empty(120),

      subsectionHead("Los platos que definen la cocina rumana"),
      bullet("Mici (o Mititei) — El alma de las parrillas rumanas. Rollitos de carne picada —cordero y ternera, sin funda— a las brasas, perfumados con ajo, tomillo y una mezcla de especias que cada carnicero guarda como secreto de familia. Se comen con mostaza dulce y una cerveza fría. En origen eran la comida del trabajador; hoy son el snack nacional que se encuentra en cada esquina y en cada celebración. Obligatorio el primer día.", true),
      bullet("Sarmale — El plato de las fiestas y los domingos. Hojas de col fermentada (o de parra en verano) rellenas con carne picada, arroz y especias, cocidas lentamente durante horas en salsa de tomate junto a costillas ahumadas. Se sirven con smântână (crema agria espesa) y pan de maíz. La influencia turca es evidente —los dolmades griegos son su primo cercano— pero el resultado es inequívocamente rumano."),
      bullet("Ciorbă — La familia de las sopas ácidas rumanas. La más conocida es la ciorbă de burtă (de callos), pero hay decenas de variedades: de verduras, de pollo, de pescado. El toque ácido —con bors, un fermento de salvado de trigo, o con jugo de col— es la firma de la gastronomía rumana. Una ciorbă caliente en una mañana fría de montaña es de las experiencias más reconfortantes del viaje."),
      bullet("Tochiturá — Solomillo de cerdo en salsa de vino tinto con cebolla y salchichas ahumadas, coronado con un huevo frito y acompañado de mămăligă. Contundente, sabroso, el plato de las cenas largas de invierno servido también en verano porque a los rumanos les gusta la contundencia en cualquier estación."),
      bullet("Mămăligă — La polenta rumana, hecha de harina de maíz cocida a fuego lento. Omnipresente en la mesa del norte del país, donde el maíz sustituyó históricamente al trigo. Se sirve con mantequilla y queso, con crema agria, como base de la tochiturá, o simplemente en trozos junto a cualquier guiso. Humilde y adictiva."),
      bullet("Papanași — El postre que se come aunque no haya sitio. Dos buñuelos de queso fresco —esponjosos por dentro, ligeramente crujientes por fuera— servidos sobre una nube de smântână y mermelada de arándanos del bosque. Dulces sin empachar, perfectos para terminar cualquier cena."),
      empty(160),

      subsectionHead("Gastronomía por regiones"),
      empty(40),

      subsectionHead("Muntenia y Oltenia — La cocina del sur"),
      body("La cocina del sur de Rumanía lleva marcada la huella de siglos de convivencia con el mundo otomano y balcánico. Las parrillas de mici son omnipresentes en Bucarest —el olor a carne a las brasas y comino es la bienvenida real a la capital—, y los sarmale de col fermentada se cocinan aquí con más ajo y más pimentón que en el norte. La ciorbă de burtă (de callos) es el plato de resaca oficial del país y cualquier cocinera de la llanura valaquia la hace mejor que cualquier restaurante de cinco estrellas."),
      empty(60),
      bullet("Ciorbă de perișoare: sopa ácida con albóndigas de carne y verduras, acidulada con zumo de col. Plato cotidiano y reconfortante."),
      bullet("Drob de miel: paté al horno de vísceras de cordero con hierbas frescas, tradicional en Semana Santa pero presente en verano en algunos restaurantes."),
      bullet("Salată de boeuf: la ensalada rusa rumana, con patata, zanahoria, mayonesa casera y pepinillo. Imprescindible en cualquier boda o fiesta familiar."),
      empty(120),

      subsectionHead("Transilvania y Maramureș — La cocina de las montañas"),
      body("La influencia centroeuropea —austro-húngara en las ciudades sajonas y magiares, centroeuropea rural en los valles— convierte la cocina de Transilvania y Maramureș en la más contundente y especiada del viaje. Los estofados de caza son comunes en los bosques de los Cárpatos, y el cerdo ahumado es la base de la despensa familiar de todo el norte del país. Las cenas en las pensiuni de Maramureș son la mejor experiencia gastronómica del viaje: comida casera cocinada con lo que hay en el huerto y el corral."),
      empty(60),
      bullet("Gulaș de mistreț — El estofado de jabalí es el plato de montaña por excelencia de Transilvania y Maramureș. Carne de caza cocinada lentamente con cebolla, pimentón húngaro, comino y vino tinto, servida sobre mămăligă o con pan de masa madre. Un plato que no existe en el sur del país y que hay que pedir cuando aparece en la carta.", true),
      bullet("Cașcaval pane: queso semi-curado rebozado y frito, servido con smântână. Entrante inevitable y adictivo en cualquier restaurante de Transilvania."),
      bullet("Kürtőskalács (cozonac de chimenea): dulce de masa enrollada en torno a un palo y cocinada sobre brasas, popular en los mercados de Transilvania. Herencia de la tradición gastronómica húngara de la región."),
      empty(120),

      subsectionHead("Bucovina y Moldova — La cocina del norte y del este"),
      body("El norte del país tiene la cocina más austera y la más honesta. La mămăligă es aquí el alimento base —no el pan—, y las sopas se acidulan con borș, un fermento de salvado de trigo que da un sabor mucho más sutil y complejo que el vinagre. La miel de los pueblos de Bucovina es extraordinaria, y los hongos del bosque —boletus, rebozuelos, setas de cardo— aparecen en las sopas y guisos de otoño pero también, conservados en vinagre o secos, durante todo el año."),
      empty(60),
      bullet("Ciorbă de fasole cu afumătură: sopa de alubias con embutido ahumado y borș. El plato más representativo de la cocina popular del norte, presente en casi todos los hogares."),
      bullet("Piftie: gelatina de cerdo especiada con ajo y pimienta, servida fría. Pariente del galantín europeo, con una versión específica moldava con pimiento rojo."),
      bullet("Cozonac: el pan dulce de las fiestas, esponjoso y enriquecido con nueces, cacao o semillas de amapola. En Bucovina lo hacen en las casas mejor que en ninguna panadería."),
      empty(160),

      infoBox("Quesos rumanos que hay que probar", [
        "Brânză de burduf — Queso de oveja intenso, curado en corteza de abeto. Sabor largo y complejo",
        "Telemea — Queso blanco salado similar al feta. Ideal con tomate y cebolleta en los desayunos",
        "Cașcaval — Queso amarillo semi-curado, normal o ahumado. Base de muchos platos gratinados",
        "Urdă — Textura suave como ricotta, ligeramente dulce. Ideal en los papanași y pastas rellenas",
      ]),
      empty(160),

      subsectionHead("Las bebidas"),
      body("La pálinca merece un párrafo aparte. Este aguardiente destilado de ciruelas, peras, albaricoques o manzanas es el lubricante social de toda Rumanía, pero especialmente del norte del país. Los artesanales caseros —ilegales técnicamente, habituales en la práctica— pueden superar el 60% de alcohol y tienen una complejidad aromática que la versión comercial no alcanza. En cualquier casa rural de Maramureș o Bucovina, el primer gesto de hospitalidad es un vasito de pálinca. Rechazarlo es una descortesía."),
      empty(80),
      bullet("Țuică — Aguardiente de ciruelas más extendido fuera del norte. Más joven y suave que la pálinca, con menos graduación. Se sirve solo, a temperatura ambiente, antes de las comidas. Es la versión más accesible para quien no está acostumbrado a los destilados caseros."),
      bullet("Cerveza — Excelente y barata. Ursus, Timișoreana (desde 1718) y Ciuc son las marcas más extendidas. Precio en bar: 1,50-2 €. En julio, imprescindible tener siempre una fría a mano."),
      bullet("Vino — Los vinos rumanos son uno de los secretos mejor guardados de Europa del Este. Las regiones de Dealu Mare (tintos de Fetească Neagră, potentes y afrutados) y Cotnari (blancos aromáticos de Grasă y Tămâioasă) producen botellas que aguantan perfectamente la comparación con vinos del sur de Francia a un tercio del precio. Pedir siempre vino de producción local en los restaurantes: suele ser mejor que las marcas de supermercado."),
      empty(120),

      subsectionHead("Bebidas sin alcohol"),
      body("Rumanía tiene una tradición de bebidas refrescantes propias que muchos viajeros descubren por casualidad y acaban buscando en cada parada."),
      empty(60),
      bullet("Apă minerală — El agua mineral rumana (Borsec, Dorna, Harghita) es extraordinaria: proviene de manantiales volcánicos en los Cárpatos y tiene un perfil mineral muy marcado. Siempre pedir 'naturală' (sin gas) o 'cu gaz' (con gas) para aclarar."),
      bullet("Limonadă — La limonada casera servida en muchos restaurantes rurales: limón, miel y agua mineral. Simple y perfecta para el calor de julio."),
      bullet("Suc de fructe — Zumos de fruta natural, especialmente de arándanos (afine) y escaramujo (măceșe). El zumo de escaramujo tiene un sabor entre ácido y floral completamente distinto a cualquier cosa conocida."),
      bullet("Socată — Refresco fermentado de flores de saúco, limón y azúcar. Casero, ligeramente burbujeante, muy popular en verano. Si aparece en la carta o en el mercado, probar sin dudarlo."),
      empty(120),

      subsectionHead("Dulces y repostería regional"),
      body("La repostería rumana es discreta pero tiene joyas que conviene buscar. No encontrarás pastelerías en cada esquina como en Austria o Francia, pero cuando aparecen los dulces tradicionales, suelen ser memorables."),
      empty(60),
      bullet("Cozonac — El pan dulce de las fiestas, esponjoso y enriquecido con nueces, cacao o semillas de amapola. En Bucovina lo hacen las abuelas mejor que ninguna panadería. Si te lo ofrecen en una casa rural, es una señal de hospitalidad de primer orden."),
      bullet("Savarina — Bizcocho empapado en almíbar de ron con nata montada. Herencia de la repostería francesa del siglo XIX que llegó a Rumanía a través de la influencia parisina de la burguesía valaca. Se encuentra en cafeterías de Bucarest y Cluj."),
      bullet("Gogoși — Donuts fritos rellenos de mermelada o crema. Omnipresentes en mercados y ferias. El olor a masa frita en el mercado de Sighetul un sábado por la mañana es una de esas memorias olfativas del viaje."),
      bullet("Plăcintă — Pastel de masa fina relleno de queso (brânză), espinacas, manzana o cereza, según la región. La versión de Maramureș (con queso y eneldo) es especialmente buena. Se vende en trozos grandes por poquísimo dinero."),
      empty(120),

      infoBox("En la ruta: dónde comer bien", [
        "Bucarest: Caru' cu Bere (cervecería histórica de 1879) y Vatra (cocina tradicional sin trampa)",
        "Brașov: Sergiana (sarmale y tochiturá definitivos, reserva recomendada en julio)",
        "Sibiu: Hermania (fusión sajona-rumana) y el mercado cubierto de Piața Mare para picoteo",
        "Maramureș: cenar en la pensiune siempre que sea posible — la comida casera es la mejor del viaje",
        "Bucovina: preguntar en el monasterio o en el pueblo por el restaurante que comen los locales",
      ]),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // ARTESANÍA Y PRODUCTOS LOCALES
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Artesanía Rumana"),
      body("Rumanía tiene una de las tradiciones artesanales más ricas y vivas de Europa. A diferencia de otros países donde la artesanía sobrevive como recreación cultural o producto turístico, en Rumanía —especialmente en el norte— sigue siendo producción real: objetos que se usan, ropa que se lleva y construcciones que se habitan. Conocer qué se hace, cómo y por qué es una de las formas más profundas de entender el país."),
      empty(120),

      subsectionHead("Madera tallada — El lenguaje visual de Maramureș"),
      body("La talla en madera es la forma de expresión artística más antigua y característica de Rumanía, y Maramureș es su epicentro. Los portones de madera de roble o haya tallados que se levantan a la entrada de cada casa son el carnet de identidad de las familias: sus motivos —el árbol de la vida, el sol, la cuerda sin fin (simbolismo dácico prerrumano)— cuentan quiénes son sus propietarios. Los maestros carpinteros aprenden el oficio de sus padres y lo transmiten a sus hijos en un ciclo que lleva siglos sin romperse."),
      empty(80),
      body("Las iglesias de madera de Maramureș son la cima de esta tradición: construidas sin un solo clavo, con técnicas medievales de ensamblaje, algunas de sus torres alcanzan 54 metros de altura. Los artesanos que las construyeron entendían la física de la madera mejor que cualquier ingeniero moderno."),
      empty(120),

      subsectionHead("Cerámica — Dos tradiciones, dos mundos"),
      body("La cerámica negra de Marginea (en Bucovina) es una de las tradiciones alfareras más antiguas de Europa, con raíces en la Edad del Hierro dácica. Su acabado negro brillante —conseguido mediante una oxidación controlada durante la cocción— y sus motivos geométricos la hacen inconfundible. No tiene color: solo negro sobre negro, con la forma y la textura como único ornamento."),
      empty(80),
      body("La cerámica de Horezu (en Valaquia, patrimonio inmaterial de la UNESCO desde 2012) representa el polo opuesto: platos y jarras de arcilla roja decorados con espirales, gallos y flores en azul, verde y marrón sobre fondo crema. Es la cerámica de fiesta, la que se saca para las celebraciones y se expone en las paredes de las casas como muestra de prosperidad."),
      empty(120),

      subsectionHead("Íconos sobre vidrio — La pintura del pueblo"),
      body("Los íconos pintados sobre vidrio son una tradición específica de Transilvania, especialmente viva en los pueblos de Sibiel y Nicula. Surgieron en los siglos XVII y XVIII como alternativa popular a los íconos pintados sobre tabla —más caros y accesibles solo para los ricos—, y su técnica es inversa: se pinta sobre la cara posterior del vidrio, de manera que el espectador ve la imagen a través de él. Los colores son planos, las figuras estilizadas y frontales, con una ingenuidad expresiva que tiene más de icono medieval que de pintura académica."),
      empty(80),
      body("El pueblo de Sibiel, a 18 km de Sibiu, alberga la colección de íconos sobre vidrio más grande de Rumanía (más de 600 piezas) en una pequeña iglesia que actúa como museo. Algunos artesanos del pueblo siguen pintando íconos con las mismas técnicas del siglo XVIII."),
      empty(120),

      subsectionHead("Huevos pintados — El arte de lo efímero"),
      body("Los huevos pintados de Pascua (ouă încondeiate) son quizás el objeto artesanal más sorprendente de Rumanía por la contradicción que encierran: una técnica de una complejidad extraordinaria aplicada a un objeto fundamentalmente frágil. La tradición, especialmente viva en el norte de Bucovina y en el pueblo de Ciocănești, consiste en aplicar cera caliente sobre la cáscara del huevo para reservar áreas del color original, teñir el resto, y repetir el proceso capa a capa hasta obtener un diseño geométrico de una precisión asombrosa."),
      empty(80),
      body("Los motivos tienen significados precisos: la espiral representa la vida eterna, el árbol de la vida la continuidad de las generaciones, el gallo el amanecer y la renovación. Un huevo bien pintado puede llevar cuatro o cinco días de trabajo. El resultado es un objeto que cabe en la palma de la mano y que concentra siglos de simbolismo."),
      empty(120),

      subsectionHead("Textiles y bordados — La ropa que cuenta historias"),
      body("La ie (pronunciar 'ie') es la blusa bordada tradicional rumana, declarada Patrimonio Cultural Inmaterial de la UNESCO. Cada región de Rumanía tiene sus propios motivos, colores y técnicas: los bordados rojos y negros sobre lino blanco de Bucovina son completamente distintos a los geométricos en negro de Maramureș o a los florales multicolores de Muntenia. En los mercados del norte, las mujeres venden ies bordadas a mano por sus propias madres y abuelas."),
      empty(80),
      body("El 24 de junio —día de La Blouse Roumaine— se celebra en toda Rumanía el Día Internacional de la Ie, en el que miles de personas llevan la blusa tradicional en señal de orgullo cultural. Si el viaje cae cerca de esa fecha, el espectáculo de ver ciudades enteras vestidas de blanco bordado es difícil de olvidar."),
      empty(120),

      infoBox("Cómo incorporar la artesanía al viaje", [
        "Sibiu (Museo ASTRA): la mejor selección curada de artesanía auténtica de todo el país. Precios justos y calidad garantizada",
        "Sighișoara (ciudadela): talleres activos dentro del recinto medieval. Joyas de plata con motivos dácicos y cuero trabajado a mano",
        "Maramureș (mercado de Sighetul, sábados): artesanía sin intermediarios. Tallas de madera, textiles y pálinca casera directamente del productor",
        "Ciocănești (Bucovina): pueblo especializado en huevos pintados. Cada casa tiene su taller",
        "Bucarest (Museo del Campesino): la mejor tienda del país para cerrar el viaje. Todo curado, todo auténtico",
      ]),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // ALOJAMIENTO
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Alojamiento"),
      body("Rumanía ofrece una gama de alojamiento que va desde hoteles de ciudad perfectamente competitivos con los estándares europeos hasta pensiuni rurales donde la anfitriona te da el desayuno en la misma mesa donde come su familia. Para un viajero en solitario que busca autenticidad, las opciones más interesantes no son siempre las más obvias."),
      empty(120),

      subsectionHead("Hoteles en ciudad — Bucarest, Brașov, Sibiu, Cluj"),
      body("En las ciudades principales la oferta hotelera es amplia y bien valorada. Un hotel de 3 estrellas bien ubicado —con parking incluido, desayuno y wifi fiable— está entre 35 y 50 euros la noche. Las cadenas internacionales están presentes, pero las opciones de boutique en los cascos históricos son más interesantes y a menudo más baratas que en la competencia centroeuropea."),
      empty(60),
      bullet("Bucarest: buscar en el barrio de Floreasca o cerca de Piața Unirii para estar bien comunicado sin pagar el premium de Lipscani."),
      bullet("Brașov: los hoteles dentro del casco histórico (zona peatonal) son perfectos para no necesitar el coche dos días."),
      bullet("Sibiu: la zona del centro histórico tiene opciones excelentes. Reservar con antelación en julio."),
      bullet("Cluj-Napoca: ciudad universitaria con buena relación calidad-precio. Muchos viajeros subestiman su oferta hotelera."),
      empty(120),

      subsectionHead("Pensiuni — La alternativa rural"),
      body("La pensiune (pensión rural) es la columna vertebral del alojamiento fuera de las ciudades. No son hostales ni apartamentos turísticos: son casas privadas con habitaciones —generalmente dos o tres— donde la familia sigue viviendo y donde la anfitriona cocina el desayuno y, si se le pide, la cena. Son la mejor experiencia del viaje."),
      empty(60),
      bullet("Maramureș: el alojamiento en pensiune es prácticamente la única opción fuera de Sighetul Marmației, y también la mejor. Precio: 20-30 € con desayuno incluido."),
      bullet("Bucovina: el circuito de monasterios tiene pensiuni bien mantenidas, algunas gestionadas directamente por comunidades monásticas."),
      bullet("Transfăgărășan: las cabañas junto a Bâlea Lac (a 2.034 m de altitud) ofrecen la experiencia más extrema del viaje. Frío garantizado incluso en julio."),
      bullet("Booking.com y Airbnb funcionan, pero las mejores pensiuni se reservan por teléfono o directamente en persona. Si llegas sin reserva y ves una señal 'Cazare' (alojamiento) en la puerta, entra y pregunta."),
      empty(120),

      subsectionHead("Campings y opciones de bajo coste"),
      body("Rumanía tiene una red de campings en crecimiento, especialmente en zonas de senderismo y en el Delta del Danubio. Para este itinerario, los campings no son la opción principal —el viaje en coche de un sitio a otro cada uno o dos días hace difícil montar y desmontar tienda—, pero pueden ser útiles en la Transfăgărășan o en los valles de Bucovina si el plan cambia sobre la marcha."),
      empty(60),
      infoBox("Consejo de reservas", [
        "Julio es temporada alta en Romania. Reservar Brașov y Maramureș con al menos tres semanas de antelación",
        "En Sibiu, la semana del festival internacional de teatro (FITS, finales de junio) puede afectar a precios de la primera semana de julio",
        "Las pensiuni de Maramureș y Bucovina suelen tener solo 2-3 habitaciones: reservar en cuanto el itinerario esté fijado",
        "Para la noche previa al vuelo de regreso desde Bucarest, reservar el hotel más cercano al aeropuerto OTP: el vuelo sale a las 06:45",
      ]),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // QUÉ RESERVAR ANTES
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Qué Reservar Antes"),
      body("Rumanía en julio es temporada alta. Algunos elementos del viaje pueden parecer secundarios y luego resultar imposibles de conseguir sin reserva previa. Esta es la lista de lo que conviene tener organizado antes de salir de España, ordenado por urgencia."),
      empty(120),

      subsectionHead("Con meses de antelación"),
      bullet("Vuelo Alicante → Bucarest (OTP) con Wizz Air — Los vuelos directos en julio tienen plazas limitadas. Reservar en cuanto el itinerario esté claro.", true),
      bullet("Coche de alquiler — El precio puede bajar un 25-30% reservando con 2-3 meses de antelación. Europcar, Avis y Budget tienen oficina en el aeropuerto OTP. Confirmar que incluye conductor único y sin restricción de kilómetros.", true),
      bullet("Alojamiento en Maramureș — Las pensiuni tienen solo 2-3 habitaciones. En julio están llenas con semanas de antelación. Reservar en cuanto el itinerario esté fijado."),
      bullet("Alojamiento en Bucovina — Similar a Maramureș. El circuito de monasterios tiene pocas plazas de calidad. Reservar pronto."),
      bullet("Hotel junto al aeropuerto OTP (última noche) — El vuelo de regreso sale a las 06:45. Necesitarás estar en el aeropuerto a las 04:30-05:00. Reservar el hotel más cercano a OTP para esa noche."),
      empty(120),

      subsectionHead("Con semanas de antelación"),
      bullet("Mocănița (Vía Estrecha del Valle de Vaser, Maramureș) — Las plazas se agotan en julio. Reservar en cffviseudesus.ro varias semanas antes.", true),
      bullet("Excursión de avistamiento de osos cerca de Brașov — Los hides tienen aforo muy limitado (8-12 personas). Reservar con 2-3 semanas de antelación."),
      bullet("Entradas al Castillo de Peleș — En julio puede haber colas de 1-2 horas sin reserva previa. Comprar online en peles.ro."),
      bullet("Tour en moto por la Transfăgărășan (si finalmente se hace) — Los operadores locales tienen pocas motos disponibles. Reservar con 2-3 semanas."),
      empty(120),

      subsectionHead("Dejar para el momento"),
      bullet("Restaurantes en ciudades — Salvo Sergiana en Brașov (recomendable reservar en julio), los demás admiten walk-in o reserva el mismo día por teléfono."),
      bullet("Museos pequeños y monasterios — Entrada directa en taquilla. Sin esperas significativas fuera de Peleș."),
      bullet("Gargantas de Bicaz y Lago Rojo — Acceso libre sin reserva."),
      bullet("Mercados locales — Sin reserva posible ni necesaria. Solo madrugad."),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // GUÍA DE SUPERVIVENCIA
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Guía de Supervivencia"),
      body("Todo lo que necesitas saber para moverte con soltura por Rumanía: desde cómo pagar la autopista hasta cómo brindar sin ofender. Estos detalles prácticos marcan la diferencia entre un viaje cómodo y uno con fricciones innecesarias."),
      empty(120),

      subsectionHead("Transporte y conducción"),
      bullet("Rovinieta (peaje electrónico) — OBLIGATORIA. Sin ella, circular por cualquier carretera nacional (DN) o autovía es ilegal. Las cámaras de control son automáticas y la multa va de 150 a 400 lei (30-80 €), sin margen de negociación. Comprarla el primer día, antes de salir del aeropuerto o de Bucarest: en www.roviniete.ro (tarjeta, inmediata), en cualquier gasolinera o en estancos. Para un coche de turismo: ~4 € (7 días) o ~7 € (30 días). No hace falta pegatina ni papel: el sistema es 100% electrónico y la matrícula queda registrada.", true),
      bullet("Waze vs Google Maps — Waze es más fiable para Rumanía: detecta mejor los radares, los baches conocidos y los atajos en carreteras rurales. Descargar mapas offline de Google Maps como respaldo para zonas sin cobertura."),
      bullet("Gasolineras — Abundantes en ciudades. Escasas en la Transfăgărășan y carreteras de montaña. Regla de oro: repostar siempre que el depósito esté por debajo del 50% al entrar en zona de montaña."),
      bullet("Conducción nocturna — Desaconsejada en zonas rurales: animales en la carretera, sin iluminación y baches imprevistos. Mejor terminar las etapas antes de anochecer."),
      bullet("Aparcamiento en ciudades — Los centros históricos tienen zonas azules de pago. Buscar aparcamientos en los bordes del casco histórico (generalmente gratuitos o con tarifa fija baja)."),
      empty(120),

      subsectionHead("Idioma — Pequeño manual de supervivencia"),
      body("El rumano es una lengua romance, pariente del español. Con un poco de atención, muchas palabras son reconocibles. En ciudades y zonas turísticas el inglés funciona bien; en zonas rurales conviene tener las frases básicas a mano."),
      empty(80),
      bullet("Bună ziua — Buenos días / Buenas tardes"),
      bullet("Bună dimineața — Buenos días (solo por la mañana)"),
      bullet("Mulțumesc — Gracias (pronunciar 'multsoomesc')"),
      bullet("Vă rog — Por favor"),
      bullet("Cât costă? — ¿Cuánto cuesta?"),
      bullet("Un bere, vă rog — Una cerveza, por favor"),
      bullet("Puteți plăti cu cardul? — ¿Se puede pagar con tarjeta?"),
      bullet("Unde pot parca? — ¿Dónde puedo aparcar?"),
      bullet("Plin, vă rog — Lleno, por favor (en la gasolinera)"),
      bullet("Nu vorbesc română — No hablo rumano"),
      bullet("Este deschis? — ¿Está abierto?"),
      bullet("A fost foarte bun, mulțumesc — Estaba muy bueno, gracias"),
      bullet("Nu înțeleg — No entiendo"),
      empty(120),

      subsectionHead("Moneda y pagos"),
      body("La moneda es el Leu rumano (RON). 1 euro equivale aproximadamente a 5 RON. Rumanía no pertenece a la zona euro."),
      empty(80),
      bullet("Efectivo imprescindible en zonas rurales — Maramureș, carreteras de montaña, mercados locales y pequeños restaurantes. Llevar siempre al menos 200-300 lei en efectivo."),
      bullet("Cajeros automáticos — Usar cajeros de bancos locales (Banca Transilvania, BRD): mejores tasas y sin comisiones ocultas."),
      bullet("Tarjeta en ciudades — Hoteles, restaurantes de ciudades y grandes supermercados aceptan Visa y Mastercard. Algunos pequeños negocios añaden recargo del 1-2%."),
      bullet("No cambiar dinero en casas de cambio de la calle — Tasas desfavorables. Mejor cajeros o recepción del hotel."),
      empty(120),

      subsectionHead("Costumbres y etiqueta"),
      bullet("Propina — 10% en restaurantes es lo correcto. En pensiuneas rurales dejar algo extra al marcharse es un gesto muy apreciado."),
      bullet("Monasterios y iglesias — Vestimenta discreta siempre: sin camisetas de tirantes, pantalones cortos o escotes. Suelen prestar pañuelos y faldas a la entrada."),
      bullet("La pálinca de bienvenida — Si en una casa rural te ofrecen un vasito al llegar, rechazarlo es una descortesía. Aceptar, agradecer y beber aunque sea un sorbo."),
      empty(120),

      subsectionHead("Osos en los Cárpatos — Cómo actuar"),
      body("Rumanía alberga la mayor población de osos pardos de Europa fuera de Rusia: entre 6.000 y 8.000 individuos, concentrados principalmente en los Cárpatos. El encuentro con un oso es posible —especialmente en las etapas de montaña— y, aunque los incidentes graves son raros, conviene conocer las normas básicas de comportamiento. Un oso que se siente sorprendido o acorralado es un oso potencialmente peligroso."),
      empty(80),
      bullet("Hacer ruido al caminar — En zonas de bosque, habla o lleva un silbato. Los osos huyen si te oyen llegar; el problema es cuando te ven de repente a poca distancia.", true),
      bullet("No acercarse nunca — Si avistas un oso a distancia, detente, no grites, no corras y retrocede despacio sin darle la espalda. Los osos pueden alcanzar 50 km/h en carretera: es imposible escapar corriendo."),
      bullet("No dejar comida en el coche ni en la tienda — El olfato del oso es siete veces más potente que el de un perro. Una bolsa con restos de comida en el maletero puede hacer que el oso dañe el coche para acceder."),
      bullet("En caso de ataque — La mayoría de los ataques son defensivos (el oso quiere que te vayas). Hazte el muerto: échate en el suelo boca abajo, protégete la nuca con las manos y permanece inmóvil hasta que el oso se aleje. Solo en ataques depredadores (muy raros) se recomienda luchar."),
      bullet("Avistamiento organizado cerca de Brașov — Varias empresas locales ofrecen salidas nocturnas a hides (observatorios de madera) en el bosque donde es habitual ver osos a 20-30 metros en total seguridad. Precio: 50-60 € por persona. Imprescindible si el viaje lo permite."),
      empty(160),
      infoBox("En la Transfăgărășan y zonas de montaña", [
        "Es común ver osos junto a la carretera de la Transfăgărășan, especialmente al atardecer y al amanecer",
        "No detenerte a darles de comer (prohibido y peligroso): un oso habituado a los humanos pierde el miedo y se vuelve imprevisible",
        "Si ves un oso desde el coche, puedes parar a fotografiarlo con el motor en marcha y las ventanillas subidas, a distancia prudente",
        "En el aparcamiento de Bâlea Lac (2.034 m), los osos son visitantes habituales: guarda absolutamente toda la comida en el maletero cerrado",
        "RO-Alert: el sistema de emergencias rumano envía avisos automáticos al móvil si se detecta presencia de osos u otras emergencias en el municipio o carretera donde te encuentras. Actívalo al cruzar la frontera (funciona igual que el ES-Alert español)",
        "App Salvamont: descarga antes de salir de España la app del servicio de rescate en montaña rumano. En caso de emergencia en la Transfăgărășan, Bicaz o cualquier zona de montaña, permite enviar tu ubicación exacta a los equipos de rescate con un solo botón",
      ]),
      empty(120),

      subsectionHead("Horarios que conviene conocer"),
      bullet("Restaurantes — Comidas: 12h-15h. Cenas: 19h-22h. En pueblos pequeños muchos cierran a las 21h."),
      bullet("Museos — La gran mayoría cierra los lunes. Muchos también hacen pausa de 13h a 14h."),
      bullet("Monasterios — Generalmente accesibles desde el amanecer hasta las 18h-20h en verano."),
      bullet("Mercados — Por las mañanas (6h-12h aprox.), generalmente los sábados. El de Sighetul Marmației es el más conocido del norte."),
      bullet("Supermercados — Kaufland, Lidl y Carrefour: 8h-22h de lunes a sábado, 9h-20h domingos."),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // PRESUPUESTO
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Presupuesto Aproximado"),
      body("El siguiente presupuesto está calculado para un viajero solo con nivel de gasto medio: casas rurales y hoteles de 3 estrellas, restaurantes locales sin lujos excesivos pero sin renunciar a las experiencias especiales. Los precios están en euros y son estimaciones para julio de 2027 con una variación esperada de ±15%."),
      empty(120),

      subsectionHead("Vuelo (aparte del presupuesto principal)"),
      body("Los vuelos Alicante (ALC) → Bucarest (OTP) con Wizz Air oscilan entre 80-200 € ida y vuelta por persona. El vuelo de ida sale a las 10:15 y llega a las 14:40 (3h 25 min); el de vuelta sale a las 06:45 desde OTP y aterriza en Alicante a las 09:30. Reservar con 3-4 meses de antelación suele dar el mejor precio."),
      empty(120),

      subsectionHead("Presupuesto en destino — 12 días"),
      budgetTable([
        budgetRow("Alojamiento (hotel 3★ en ciudades / pensiune en rural)", "30-45 €", "360-540 €", true),
        budgetRow("Noche aeropuerto OTP — última noche antes del vuelo (pago único)", "—", "40-55 €", false),
        budgetRow("Coche de alquiler (categoría B, seguro incluido)", "30-40 €", "360-480 €", true),
        budgetRow("Gasolina (estimado ~200 km/día, ~7L/100km)", "15-20 €", "180-240 €", false),
        budgetRow("Rovigneta (peaje electrónico, pago único)", "—", "7-10 €", true),
        budgetRow("Comidas (desayuno + comida + cena)", "25-35 €", "300-420 €", false),
        budgetRow("Entradas a monumentos y atracciones", "10-15 €", "120-180 €", true),
        budgetRow("Excursión avistamiento de osos — Brașov (pago único)", "—", "50-60 €", false),
        budgetRow("Tour en moto Transfăgărășan — opcional (pago único)", "—", "~160 €", true),
        budgetRow("Artesanía, souvenirs, propinas e imprevistos", "5-10 €", "60-120 €", false),
      ]),
      empty(160),

      highlight("Total estimado sin vuelo ni moto opcional: ~1.437-2.050 €. Con tour en moto: ~1.597-2.210 €. Estimación central: ~1.700 € en destino + ~120 € en vuelo = ~1.820 € por persona para 12 días."),
      empty(120),

      infoBox("Consejos para ahorrar sin sacrificar calidad", [
        "Reservar el coche con 2-3 meses de antelación: el precio puede bajar un 25-30%",
        "Las pensiuneas rurales de Maramureș y Bucovina son más baratas y auténticas que los hoteles de ciudad",
        "El menú del mediodía en un restaurante local cuesta 5-8 € con sopa incluida: el mejor ratio calidad-precio",
        "Usar efectivo: muchos pequeños restaurantes y comercios aplican recargo por tarjeta",
        "Los supermercados Kaufland o Lidl son perfectos para desayunos y picnics de carretera",
        "Los monasterios de Bucovina cuestan solo 2 € por entrada: el mejor valor cultural del viaje",
      ]),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // IMPRESCINDIBLE EN LA MALETA
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Imprescindible en la Maleta"),
      body("Rumanía en julio combina calor en ciudades, frío en altura y humedad en zonas de montaña. El equipaje ideal es versátil, compacto y pensado para el coche."),
      empty(120),

      infoBox("Ropa y calzado", [
        "Ropa en capas: la temperatura en la Transfăgărășan (2.000 m) puede ser 20 °C más fría que en Bucarest el mismo día",
        "Calzado de senderismo con suela antideslizante: imprescindible para las gargantas de Bicaz y senderos de montaña",
        "Ropa discreta para monasterios: camisas de manga larga y pantalón largo (al menos uno por persona)",
        "Sandalia cómoda para ciudades: el adoquinado de Sighișoara y Sibiu agradece suela blanda",
        "Impermeable ligero o chubasquero: los chubascos de montaña en julio son rápidos pero intensos",
      ]),
      empty(120),

      infoBox("Documentación y dinero", [
        "DNI español en vigor — suficiente como ciudadano de la UE, sin pasaporte ni visado",
        "Tarjeta sanitaria europea (TSJE) más seguro de viaje con asistencia médica",
        "Efectivo en leu rumano (RON): llevar 300-400 RON desde el primer momento",
        "Copia digital de todos los documentos: foto en el móvil de DNI, tarjetas y reservas",
        "No se necesita adaptador de enchufe: los enchufes rumanos son tipo C y F, iguales que en España",
      ]),
      empty(120),

      infoBox("El kit de supervivencia del viajero en Rumanía", [
        "Repelente de insectos — Imprescindible en zonas de bosque (Maramureș) y junto a lagos. Los mosquitos del norte pueden ser agresivos en julio",
        "Botella de agua reutilizable — El agua del grifo es potable en ciudades. Llevar siempre reserva para los tramos de montaña",
        "Bolsa plegable reutilizable — Para mercados, el Museo ASTRA y guardar compras de artesanía. Los plásticos de un solo uso están limitados",
        "Mapas offline descargados — Google Maps offline de todas las regiones. En los Cárpatos y Maramureș la cobertura 4G puede fallar",
        "Powerbank de al menos 10.000 mAh — Los días de conducción con GPS activo y música agotan la batería del móvil",
        "Linterna de mano o frontal — Para visitas a minas (Salina Turda), grutas y monasterios con poca iluminación interior",
        "Bolsa de medicamentos básicos — Ibuprofeno, antihistamínico, tiritas y crema antiséptica",
      ]),

      // ═══════════════════════════════════════════════════════
      // ERRORES QUE YO EVITARÍA
      // ═══════════════════════════════════════════════════════
      pageBreakPara(),
      sectionHead("Errores que Yo Evitaría"),
      body("Todo viajero que regresa de Rumanía tiene una lista de cosas que haría diferente. Esta página recoge los errores más comunes para que no llegues a ellos por tu cuenta."),
      empty(120),

      subsectionHead("En la carretera"),
      bullet("Conducir de noche — Las carreteras rurales de Rumanía tienen animales sueltos (vacas, caballos, perros), baches sin señalizar e iluminación inexistente. Un día que se alarga hasta el anochecer se convierte en un riesgo real. Planificar siempre llegar antes del atardecer.", true),
      bullet("No activar la Rovinieta el primer día — La cámara de control está a la salida del aeropuerto. La multa llega al domicilio semanas después. Comprarla antes de arrancar el coche, sin excepción."),
      bullet("Subestimar la Transfăgărășan — El GPS dice 133 km pero la carretera tarda 3h 30 min en condiciones normales. En julio hay colas en los miradores, rebaños cruzando la carretera y paradas inevitables. Salir pronto y no planificar nada en el destino antes de las 18h."),
      bullet("Planificar demasiados kilómetros en un día — En Rumanía, 200 km por carreteras de montaña equivalen a 4 horas de conducción real. Los días con más de 200 km no deberían tener agenda densa al llegar."),
      empty(120),

      subsectionHead("En los pueblos y ciudades"),
      bullet("Cambiar dinero en el aeropuerto — Las casas de cambio del aeropuerto OTP aplican tasas pésimas. Sacar lei directamente de un cajero del banco Banca Transilvania dentro del aeropuerto o en la primera ciudad."),
      bullet("Dar de comer a los osos — Un oso acostumbrado a recibir comida de los coches de la Transfăgărășan pierde el miedo a los humanos y se vuelve peligroso. Está prohibido y es una irresponsabilidad con los siguientes viajeros."),
      bullet("Entrar a los monasterios sin ropa adecuada — No importa el calor de julio: pantalón largo y mangas en los monasterios de Bucovina y Transilvania. Los guardianes piden salir o prestan pañuelos, pero la primera impresión se pierde."),
      bullet("Ignorar los precios antes de sentarte en un restaurante turístico — En Sighișoara y Brașov (ciudadela) hay restaurantes con carta en varios idiomas y precios muy distintos para turistas. Mirar la carta y los precios antes de sentarse."),
      empty(120),

      subsectionHead("En la planificación"),
      bullet("Reservar la Mocănița el mismo día — En julio las plazas del tren de vapor del Valle de Vaser se agotan semanas antes. Ver sección 'Qué reservar antes'."),
      bullet("No descargar mapas offline — En los Cárpatos y Maramureș la cobertura 4G falla en los valles. Descargar Google Maps offline de todas las regiones antes de salir de España."),
      bullet("Subestimar Maramureș — Es la región que más sorprende a todos los viajeros que la visiten. No tratarla como tránsito hacia Bucovina: merece al menos día y medio de ritmo lento."),
      empty(200),

      // ═══════════════════════════════════════════════════════
      // FINAL NOTE
      // ═══════════════════════════════════════════════════════
      new Table({
        columnWidths: [9026],
        width: { size: 9026, type: WidthType.DXA },
        borders: NB,
        rows: [new TableRow({ children: [new TableCell({
          width: { size: 9026, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: C.navy },
          margins: { top: 200, bottom: 200, left: 400, right: 400 },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
              children: [new TextRun({ text: "★  Bine ați venit în România  ★", color: C.goldLight, bold: true, size: 28, font: "Georgia", characterSpacing: 60 })]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
              children: [new TextRun({ text: "(Bienvenido a Rumanía)", color: "B0BEC5", size: 22, font: "Calibri", italics: true })]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER, spacing: { before: 80, after: 0 },
              children: [new TextRun({ text: "Guía preparada para Pablo Lloret · Julio 2027 · El Secreto de los Cárpatos", color: "78909C", size: 20, font: "Calibri" })]
            }),
          ]
        })]})],
      }),
      empty(80),
    ]
  }
];

const doc = new Document({ sections });

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('./rumania_guia_viaje.docx', buffer);
  console.log('Documento generado: ./rumania_guia_viaje.docx');

  // Exportación automática a PDF con LibreOffice, para poder revisar el
  // maquetado sin pasos manuales. Es opcional: si falla, el .docx ya está
  // escrito y sigue siendo válido.
  try {
    execSync('soffice --headless --convert-to pdf ./rumania_guia_viaje.docx', { stdio: 'ignore' });
    console.log('PDF generado: ./rumania_guia_viaje.pdf');
  } catch (e) {
    console.warn('Aviso: no se pudo generar el PDF (¿LibreOffice instalado?):', e.message);
  }
});
