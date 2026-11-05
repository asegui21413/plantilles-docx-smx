const { Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType, 
        BorderStyle, PageNumber, TabStopType, TabStopPosition } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Calibri", size: 22 }
      }
    },
    paragraphStyles: [
      {
        id: "Title",
        name: "Title",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 36, bold: true, font: "Georgia", color: "1B4F72" },
        paragraph: { spacing: { before: 200, after: 200 }, alignment: AlignmentType.CENTER }
      },
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 28, bold: true, font: "Calibri", color: "1B4F72" },
        paragraph: { spacing: { before: 240, after: 120 } }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } // 2cm margins
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "CLUB MARÍTIM DE MAÓ", bold: true, size: 20, font: "Georgia", color: "1B4F72" })
            ],
            border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "1B4F72", space: 8 } }
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60 },
            children: [
              new TextRun({ text: "Carrer de Sant Roc, 15 · 07701 Maó (Menorca) · Tel. 971 35 XX XX · info@clubmaritimmao.cat", size: 16, color: "666666", font: "Calibri" })
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 8 } },
            children: [
              new TextRun({ text: "Pàgina ", size: 16, color: "666666" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "666666" }),
              new TextRun({ text: " | Document oficial del Club Marítim de Maó", size: 16, color: "666666" })
            ]
          })
        ]
      })
    },
    children: [
      new Paragraph({ spacing: { before: 400 }, children: [] }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 },
        children: [
          new TextRun({ text: "CONVOCATÒRIA DE REUNIÓ", bold: true, size: 36, font: "Georgia", color: "1B4F72" })
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: "2874A6", space: 4 } },
        children: [
          new TextRun({ text: "Assemblea General Ordinària", size: 26, italics: true, color: "2874A6", font: "Calibri" })
        ]
      }),
      new Paragraph({
        spacing: { before: 300, after: 200 },
        children: [
          new TextRun({ text: "Maó, 5 de novembre de 2026", size: 22 })
        ]
      }),
      new Paragraph({
        spacing: { before: 200, after: 200 },
        children: [
          new TextRun({ text: "Benvolguts/des socis i sòcies,", size: 22 })
        ]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [
          new TextRun({ text: "Per mitjà de la present, el ", size: 22 }),
          new TextRun({ text: "Club Marítim de Maó", bold: true, size: 22 }),
          new TextRun({ text: " us convoca a la propera ", size: 22 }),
          new TextRun({ text: "Assemblea General Ordinària", bold: true, size: 22 }),
          new TextRun({ text: " que tindrà lloc el proper dilluns a les ", size: 22 }),
          new TextRun({ text: "20:00 hores", bold: true, size: 22, color: "C0392B" }),
          new TextRun({ text: " a la ", size: 22 }),
          new TextRun({ text: "Sala d'Actes del Club", bold: true, size: 22 }),
          new TextRun({ text: ".", size: 22 })
        ]
      }),
      new Paragraph({
        spacing: { before: 300, after: 120 },
        children: [
          new TextRun({ text: "Ordre del dia:", bold: true, size: 24, color: "1B4F72" })
        ]
      }),
      new Paragraph({
        spacing: { after: 80 },
        indent: { left: 360 },
        children: [new TextRun({ text: "1. Lectura i aprovació de l'acta de la sessió anterior.", size: 22 })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        indent: { left: 360 },
        children: [new TextRun({ text: "2. Informe de la Junta Directiva sobre l'estat del Club.", size: 22 })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        indent: { left: 360 },
        children: [new TextRun({ text: "3. Presentació i aprovació dels comptes de l'exercici.", size: 22 })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        indent: { left: 360 },
        children: [new TextRun({ text: "4. Proposta de noves activitats nàutiques per a la temporada.", size: 22 })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        indent: { left: 360 },
        children: [new TextRun({ text: "5. Torn obert de paraules.", size: 22 })]
      }),
      new Paragraph({
        spacing: { before: 300, after: 200 },
        children: [
          new TextRun({ text: "Es prega puntualitat. La vostra assistència és molt important per al bon funcionament del Club.", size: 22 })
        ]
      }),
      new Paragraph({
        spacing: { before: 200, after: 200 },
        children: [
          new TextRun({ text: "Atentament,", size: 22 })
        ]
      }),
      new Paragraph({
        spacing: { before: 400 },
        children: [
          new TextRun({ text: "La Junta Directiva", bold: true, size: 22 })
        ]
      }),
      new Paragraph({
        children: [
          new TextRun({ text: "Club Marítim de Maó", size: 20, color: "666666" })
        ]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/workdir/artifacts/plantilles_docx/01_Convocatoria_Reunio.docx", buffer);
  console.log("Convocatòria creada correctament");
});
