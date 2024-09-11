import express, { Request, Response } from 'express';
import { xlsxToPdf } from './file/xlsxToPdf';
import path from 'path';
const app = express();
const port = 3000;

app.listen(port, async () => {
  console.log(`Server is running at http://localhost:${port}`);
  const xlsxFile = path.join(__dirname, '../files/test2.xlsx');
  const pdfFile = path.join(__dirname, '../files/test.pdf');

  await xlsxToPdf(xlsxFile, pdfFile)
    .then(() => {
      console.log('PDF gerado com sucesso.');
    })
    .catch((error) => {
      console.error('Erro ao converter:', error);
    });
});