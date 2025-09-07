const express = require('express');
const multer = require('multer');
const vision = require('@google-cloud/vision');
const xlsx = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors({
  origin: ['https://image-ai-tau.vercel.app', 'http://localhost:3000', 'http://localhost:3001'],
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use((req, res, next) => {
  console.log(`${req.method} ${req.path}`, req.body ? 'with body' : 'no body');
  next();
});
const upload = multer({ storage: multer.memoryStorage() });

// Initialize Google Cloud Vision client with environment variable authentication
const client = new vision.ImageAnnotatorClient({
  credentials: process.env.GOOGLE_APPLICATION_CREDENTIALS 
    ? JSON.parse(process.env.GOOGLE_APPLICATION_CREDENTIALS)
    : undefined
});

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
      res.json({ result: Buffer.from(extractedText).toString('base64') });
    } else if (req.body.type === 'excel') {
      const workbook = xlsx.utils.book_new();
      const lines = extractedText.split('\n').filter(line => line.trim());
      const data = [['Extracted Text'], ...lines.map(line => [line])];
      const worksheet = xlsx.utils.aoa_to_sheet(data);
      xlsx.utils.book_append_sheet(workbook, worksheet, 'ExtractedText');
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

app.post('/api/batch-convert', upload.array('images', 10), async (req, res) => {
  console.log('Batch conversion request received');
  
  if (!req.files || req.files.length === 0) {
    console.log('No files uploaded');
    return res.status(400).json({ error: 'No image files provided' });
  }

  console.log(`Processing ${req.files.length} files`);

  try {
    const results = [];
    
    for (let i = 0; i < req.files.length; i++) {
      const file = req.files[i];
      console.log(`Processing file ${i + 1}/${req.files.length}: ${file.originalname}`);
      
      try {
        const [result] = await client.textDetection(file.buffer);
        
        if (result.textAnnotations && result.textAnnotations.length > 0) {
          const extractedText = result.textAnnotations[0].description;
          results.push({
            filename: file.originalname,
            success: true,
            text: extractedText
          });
        } else {
          results.push({
            filename: file.originalname,
            success: false,
            error: 'No text detected in image'
          });
        }
      } catch (error) {
        console.error(`Error processing ${file.originalname}:`, error);
        results.push({
          filename: file.originalname,
          success: false,
          error: error.message
        });
      }
    }

    console.log(`Batch processing completed: ${results.filter(r => r.success).length}/${results.length} successful`);
    res.json({ results });
  } catch (error) {
    console.error('Batch conversion error:', error);
    res.status(500).json({ error: 'Batch conversion failed', details: error.message });
  }
});

app.post('/api/download', express.json(), (req, res) => {
  const { data, format } = req.body;
  console.log('Received download request for format:', format);

  if (format === 'txt') {
    try {
      const decodedText = Buffer.from(data, 'base64').toString('utf-8');
      res.setHeader('Content-Type', 'text/plain; charset=utf-8');
      res.setHeader('Content-Disposition', 'attachment; filename=converted.txt');
      res.send(decodedText);
    } catch (error) {
      console.error('Error processing text data:', error);
      res.status(500).json({ error: 'Failed to process text data' });
    }
  } else if (format === 'xlsx') {
    try {
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

// Export the Express app for Vercel
module.exports = app;

// Keep the local development server setup for development mode
if (process.env.NODE_ENV !== 'production') {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
  });
}
