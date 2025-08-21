// server.js - DOCX Template Rendering Microservice
const express = require('express');
const multer = require('multer');
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');

const app = express();
const upload = multer();

// Middleware
app.use(express.json());

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', service: 'docx-renderer' });
});

// Main DOCX rendering endpoint
app.post('/render-docx', upload.single('template'), (req, res) => {
  try {
    // Get template file from multipart upload
    const templateBuffer = req.file.buffer;
    
    // Get JSON data from request body or form field
    const templateData = typeof req.body.data === 'string' 
      ? JSON.parse(req.body.data) 
      : req.body.data || req.body;

    console.log('Processing template with data:', templateData);

    // Load the DOCX template
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, {
      errorOnMissingTags: false,
      paragraphLoop: true,
      linebreaks: true,
      syntax: {
        allowUnopenedTag: true,
        allowUnclosedTag: true,
      },
    });

    // Render the document with the provided data
    doc.render(templateData);

    // Generate the output buffer
    const buffer = doc.getZip().generate({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    });

    // Set headers for file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="rendered-document.docx"');
    res.setHeader('Content-Length', buffer.length);

    // Send the rendered DOCX
    res.send(buffer);

  } catch (error) {
    console.error('Error rendering DOCX:', error);
    res.status(500).json({ 
      error: 'Failed to render document', 
      details: error.message 
    });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`DOCX Renderer service running on port ${PORT}`);
});
