const { Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType, 
        BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType,
        VerticalAlign } = require('docx');
const fs = require('fs');

const borderNone = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const bordersNone = { top: borderNone, bottom: borderNone, left: borderNone, right: borderNone };
const thinBorder = { style: BorderStyle.SINGLE, size: 8, color: "2C3E50" };
const bottomBorder = { top: borderNone, bottom: thinBorder, left: borderNone, right: borderNone };

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Calibri", size: 20 }
      }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 720, right: 720, bottom: 720, left: 720 }
      }
    },
    children: [
      // Header with name and title
      new Table({
        width: { size: 10466, type: WidthType.DXA },
        columnWidths: [7000, 3466],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders: bordersNone,
                width: { size: 7000, type: WidthType.DXA },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: "JOAN SUREDA PONS", bold: true, size: 40, font: "Arial", color: "1A5276" })
                    ]
                  }),
                  new Paragraph({
                    spacing: { before: 60 },
                    children: [
                      new TextRun({ text: "Tècnic en Sistemes Microinformàtics i Xarxes", size: 22, color: "2874A6", font: "Calibri" })
                    ]
                  })
                ]
              }),
              new TableCell({
                borders: bordersNone,
                width: { size: 3466, type: WidthType.DXA },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "Carrer Major, 12", size: 18, color: "555555" })]
                  }),
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "07701 Maó (Menorca)", size: 18, color: "555555" })]
                  }),
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "971 35 XX XX", size: 18, color: "555555" })]
                  }),
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "joan.sureda@email.cat", size: 18, color: "555555" })]
                  })
                ]
              })
            ]
          })
        ]
      }),
      // Separator line
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: "1A5276", space: 1 } },
        spacing: { after: 200 },
        children: []
      }),
      // Profile
      new Paragraph({
        spacing: { before: 120, after: 80 },
        children: [
          new TextRun({ text: "PERFIL PROFESSIONAL", bold: true, size: 22, color: "1A5276", font: "Arial" })
        ]
      }),
      new Paragraph({
        spacing: { after: 160 },
        children: [
          new TextRun({ text: "Tècnic en SMX amb experiència en manteniment d'equips, configuració de xarxes i suport a usuaris. Orientat a la resolució de problemes, amb capacitat de treball en equip i ganes d'aprendre noves tecnologies. Disponibilitat immediata i mobilitat geogràfica.", size: 20 })
        ]
      }),
      // Experience
      new Paragraph({
        spacing: { before: 160, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "AED6F1", space: 4 } },
        children: [
          new TextRun({ text: "EXPERIÈNCIA LABORAL", bold: true, size: 22, color: "1A5276", font: "Arial" })
        ]
      }),
      new Paragraph({
        spacing: { before: 100 },
        children: [
          new TextRun({ text: "Tècnic de suport informàtic", bold: true, size: 21 }),
          new TextRun({ text: "  |  Informàtica Menorca SL", size: 20, color: "555555" })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "Juny 2025 – Actualitat  ·  Maó", size: 18, italics: true, color: "666666" })
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        indent: { left: 200 },
        children: [new TextRun({ text: "• Manteniment preventiu i correctiu d'equips informàtics", size: 19 })]
      }),
      new Paragraph({
        spacing: { after: 40 },
        indent: { left: 200 },
        children: [new TextRun({ text: "• Configuració de xarxes locals i resolució d'incidències", size: 19 })]
      }),
      new Paragraph({
        spacing: { after: 120 },
        indent: { left: 200 },
        children: [new TextRun({ text: "• Atenció a usuaris i documentació de procediments", size: 19 })]
      }),
      new Paragraph({
        spacing: { before: 80 },
        children: [
          new TextRun({ text: "Pràctiques en empresa", bold: true, size: 21 }),
          new TextRun({ text: "  |  Ajuntament de Maó", size: 20, color: "555555" })
        ]
      }),
      new Paragraph({
        spacing: { after: 60 },
        children: [
          new TextRun({ text: "Març 2025 – Maig 2025  ·  Maó", size: 18, italics: true, color: "666666" })
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        indent: { left: 200 },
        children: [new TextRun({ text: "• Suport a l'àrea d'informàtica municipal", size: 19 })]
      }),
      new Paragraph({
        spacing: { after: 160 },
        indent: { left: 200 },
        children: [new TextRun({ text: "• Inventari d'equips i actualització de software", size: 19 })]
      }),
      // Education
      new Paragraph({
        spacing: { before: 80, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "AED6F1", space: 4 } },
        children: [
          new TextRun({ text: "FORMACIÓ", bold: true, size: 22, color: "1A5276", font: "Arial" })
        ]
      }),
      new Paragraph({
        spacing: { before: 80 },
        children: [
          new TextRun({ text: "Cicle Formatiu de Grau Mitjà – SMX", bold: true, size: 21 })
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: "IES Joan Ramis i Ramis  ·  2024 – 2026", size: 19, color: "555555" })
        ]
      }),
      new Paragraph({
        spacing: { after: 120 },
        children: [
          new TextRun({ text: "Nota mitjana: 9,2  |  Premi al millor alumne de 1r de SMX", size: 19, color: "1A5276" })
        ]
      }),
      // Skills
      new Paragraph({
        spacing: { before: 80, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "AED6F1", space: 4 } },
        children: [
          new TextRun({ text: "COMPETÈNCIES TÈCNIQUES", bold: true, size: 22, color: "1A5276", font: "Arial" })
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: "Sistemes operatius: ", bold: true, size: 19 }),
          new TextRun({ text: "Windows 10/11, Linux (Ubuntu, Debian), macOS", size: 19 })
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: "Xarxes: ", bold: true, size: 19 }),
          new TextRun({ text: "TCP/IP, VLAN, Wi-Fi, firewalls, switchs gestionables", size: 19 })
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: "Ofimàtica: ", bold: true, size: 19 }),
          new TextRun({ text: "Microsoft Office, LibreOffice, Google Workspace, WPS Office", size: 19 })
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: "Altres: ", bold: true, size: 19 }),
          new TextRun({ text: "Muntatge d'equips, diagnòstic de hardware, virtualització (VirtualBox)", size: 19 })
        ]
      }),
      new Paragraph({
        spacing: { after: 160 },
        children: [
          new TextRun({ text: "Idiomes: ", bold: true, size: 19 }),
          new TextRun({ text: "Català (nadiu), Castellà (nadiu), Anglès (B1)", size: 19 })
        ]
      }),
      // Footer note
      new Paragraph({
        spacing: { before: 200 },
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Disponibilitat immediata · Referències disponibles sota demanda", size: 16, italics: true, color: "888888" })
        ]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/workdir/artifacts/plantilles_docx/02_Curriculum_Vitae.docx", buffer);
  console.log("CV creat correctament");
});
