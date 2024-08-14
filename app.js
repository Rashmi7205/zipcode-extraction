import express from 'express';
import multer from 'multer';
import XLSX from 'xlsx';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = 3000;

app.use(cors());

const upload = multer({ dest: 'uploads/' });

app.use(express.static(path.join(__dirname, 'public')));

app.post('/upload', upload.single('file'), (req, res) => {
  const filePath = req.file.path;
  const fileName = req.file.originalname;

  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  for (let i = 1; i < data.length; i++) {
    let mailingAddress = data[i][0];
    if (mailingAddress) {
      let zipCodeMatch = mailingAddress.match(/\b\d{6}\b/);
      if (zipCodeMatch) {
        data[i][2] = zipCodeMatch[0];
      }
    }
  }

  const newWorksheet = XLSX.utils.aoa_to_sheet(data);
  workbook.Sheets[sheetName] = newWorksheet;

  const newFileName = `updated_${fileName}`;
  const newFilePath = path.join(__dirname, 'uploads', newFileName);
  XLSX.writeFile(workbook, newFilePath);


  res.download(newFilePath, newFileName, (err) => {
    if (err) {
      console.error('File download error:', err);
    }
    fs.unlinkSync(filePath);
    fs.unlinkSync(newFilePath);
  });
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
