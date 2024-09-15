const express = require('express');
const multer = require('multer');
const fs = require('fs');
const mammoth = require('mammoth'); // For Word-to-HTML conversion
const ExcelJS = require('exceljs');

// Set up Express
const app = express();
const port = 3000;

// Multer configuration for file uploads
const upload = multer({ dest: 'uploads/' });

// Route to serve upload form (for testing in a browser)
app.get('/', (req, res) => {
  res.send(`
    <h2>Upload a Word file</h2>
    <form enctype="multipart/form-data" action="/upload" method="POST">
      <input type="file" name="file" />
      <button type="submit">Upload</button>
    </form>
  `);
});

// Route to handle file upload and conversion
app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path; // Path of the uploaded file

    // Step 1: Extract content from the Word file as HTML
    const result = await mammoth.convertToHtml({ path: filePath });
    const htmlContent = result.value; // Extracted HTML content

    // Step 2: Parse the content and extract explanations
    const explanations = extractExplanationsFromHTML(htmlContent);

    // Step 3: Convert explanations to HTML and write to Excel
    const excelBuffer = await convertToExcel(explanations);

    // Send Excel file as a response
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="explanations.xlsx"'
    );
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(excelBuffer);

    // Clean up the uploaded file
    fs.unlinkSync(filePath);
  } catch (error) {
    res.status(500).send('Error processing the file.');
    console.error(error);
  }
});

// Start server
app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});

// Helper function to extract explanations from HTML
function extractExplanationsFromHTML(html) {
  const explanations = [];
  const explanationRegex = /Explanation:(.*?)<\/p>/gs; // Regex to find "Explanation:" and capture the HTML content after it
  
  let match;
  while ((match = explanationRegex.exec(html)) !== null) {
    // Extract the full explanation HTML (preserving tables, bullet points, etc.)
    const explanationHTML = match[0].trim();
    explanations.push(explanationHTML);
  }

  return explanations;
}

// Helper function to convert explanations to HTML and write to Excel
async function convertToExcel(explanations) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Explanations');

  explanations.forEach((explanation, index) => {
    // Add the explanation as is (with full HTML content) to the Excel file
    sheet.addRow([explanation]);
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}
