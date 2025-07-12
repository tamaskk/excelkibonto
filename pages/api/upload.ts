import { NextApiRequest, NextApiResponse } from 'next';
import formidable from 'formidable';
import * as XLSX from 'xlsx';
import fs from 'fs';

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const form = formidable({
      uploadDir: '/tmp',
      keepExtensions: true,
      maxFileSize: 10 * 1024 * 1024, // 10MB limit
    });

    const [fields, files] = await form.parse(req);
    const file = files.file?.[0];

    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // Read the Excel file
    const workbook = XLSX.readFile(file.filepath);
    const sheets: any[] = [];

    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      sheets.push({
        sheetName,
        data: jsonData,
        rows: jsonData.length,
        columns: jsonData[0]?.length || 0,
      });
    });

    // Clean up the temporary file
    fs.unlinkSync(file.filepath);

    res.status(200).json({
      success: true,
      fileName: file.originalFilename,
      sheets,
    });
  } catch (error) {
    console.error('Error processing file:', error);
    res.status(500).json({ error: 'Error processing file' });
  }
} 