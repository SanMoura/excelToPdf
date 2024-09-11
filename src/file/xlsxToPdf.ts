import ExcelJS from 'exceljs';
import fs from 'fs';
import PdfPrinter from 'pdfmake';

// Função para ler o arquivo .xlsx
async function readXlsxFile(filePath: string): Promise<any[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const rows: any[] = [];

  workbook.eachSheet((sheet) => {
    sheet.eachRow((row) => {
      const rowData: any = [];
      row.eachCell((cell) => {
        rowData.push(cell.value);
      });
      rows.push(rowData);
    });
  });

  return rows;
}

// Função para criar e salvar o PDF usando pdfmake
async function createPdfFromXlsx(data: any[], outputPath: string): Promise<void> {
  // Definir as fontes para o pdfmake
  const fonts = {
    Roboto: {
      normal: 'Helvetica',
      bold: 'Helvetica-Bold',
      italics: 'Helvetica-Oblique',
      bolditalics: 'Helvetica-BoldOblique',
    }
  };


  const printer = new PdfPrinter(fonts);

  // Estrutura do documento PDF
  const docDefinition = {
    content: [
      {
        table: {
          body: data
        }
      }
    ]
  };

  // Criar o PDF
  const pdfDoc = printer.createPdfKitDocument(docDefinition);
  const stream = fs.createWriteStream(outputPath);
  pdfDoc.pipe(stream);
  pdfDoc.end();

  return new Promise((resolve, reject) => {
    stream.on('finish', () => resolve());
    stream.on('error', reject);
  });
}

// Função principal para converter .xlsx para PDF
export async function xlsxToPdf(xlsxFilePath: string, pdfFilePath: string): Promise<void> {
  const data = await readXlsxFile(xlsxFilePath);
  await createPdfFromXlsx(data, pdfFilePath);
}