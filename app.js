// server.js
const express = require('express');
const multer = require('multer');
const vision = require('@google-cloud/vision');
const xlsx = require('xlsx');
const XLSX = require('xlsx');
const cors = require('cors');
const { promisify } = require('util');

const app = express();
app.use(cors());
const upload = multer({ storage: multer.memoryStorage() });

// Path to your service account key file
// const keyFilePath = path.join(__dirname, 'imageai-437222-140b791e96a1.json');
const client = new vision.ImageAnnotatorClient();

app.post('/api/convert', upload.single('image'), async (req, res) => {
  console.log('Request received for conversion');
  
  if (!req.file) {
    console.log('No file uploaded');
    return res.status(400).json({ error: 'No image file provided' });
  }

  console.log(`File received: ${req.file.originalname}, Size: ${req.file.size} bytes`);

  try {
    const [result] = await client.textDetection(req.file.buffer);
    console.log('Vision API response received');

    if (!result.textAnnotations || result.textAnnotations.length === 0) {
      console.log('No text detected in the image');
      return res.status(400).json({ error: 'No text detected in the image' });
    }

    const extractedText = result.textAnnotations[0].description;
    console.log('Extracted text:', extractedText.substring(0, 100) + '...');

    if (req.body.type === 'text') {
      res.json({ result: extractedText });
    } else if (req.body.type === 'excel') {
      const workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.aoa_to_sheet([extractedText.split('\n').map(line => [line])]);
      xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
      const excelBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
      res.json({ result: excelBuffer.toString('base64') });
    } else {
      res.status(400).json({ error: 'Invalid conversion type' });
    }
  } catch (error) {
    console.error('Conversion error:', error);
    res.status(500).json({ error: 'Conversion failed', details: error.message });
  }
});

app.post('/api/download', express.json(), (req, res) => {
  const { data, format } = req.body;
  console.log('Received download request for format:', format);
  console.log('Data preview:', data.substring(0, 100) + '...');

  if (format === 'txt') {
    try {
      // Decode base64 and extract text content
      const workbook = XLSX.read(Buffer.from(data, 'base64'), { type: 'buffer' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const textContent = XLSX.utils.sheet_to_txt(worksheet);

      res.setHeader('Content-Type', 'text/plain');
      res.setHeader('Content-Disposition', 'attachment; filename=converted.txt');
      res.send(textContent);
    } catch (error) {
      console.error('Error extracting text from Excel:', error);
      res.status(500).json({ error: 'Failed to extract text content' });
    }
  } else if (format === 'xlsx') {
    try {
      // Decode base64 and send as Excel file
      const excelBuffer = Buffer.from(data, 'base64');
      console.log('Excel buffer created, size:', excelBuffer.length, 'bytes');

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename=converted.xlsx');
      res.send(excelBuffer);
    } catch (error) {
      console.error('Error processing Excel data:', error);
      res.status(500).json({ error: 'Failed to process Excel data' });
    }
  } else {
    res.status(400).json({ error: 'Invalid format' });
  }
});

// Start the server on port 3000 or the port from environment variables
const port = process.env.PORT || 3000;
app.listen(port, '0.0.0.0', () => {
  console.log(`Server running on port ${port}`);
});
