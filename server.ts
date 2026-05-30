/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from 'express';
import path from 'path';
import fs from 'fs/promises';
import { createServer as createViteServer } from 'vite';
// @ts-ignore
import libre from 'libreoffice-convert';

const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT) : 3000;
const DB_FILE_PATH = path.join(process.cwd(), 'data', 'templates_db.json');

// Ensure DB directory and file exist
async function initDB() {
  try {
    await fs.mkdir(path.dirname(DB_FILE_PATH), { recursive: true });
    try {
      await fs.access(DB_FILE_PATH);
    } catch {
      await fs.writeFile(DB_FILE_PATH, JSON.stringify([], null, 2), 'utf-8');
    }
  } catch (error) {
    console.error('Gagal inisialisasi database server:', error);
  }
}

// Support large payloads (since docx files can be encoded as large base64 strings)
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ limit: '100mb', extended: true }));

// GET /api/templates
app.get('/api/templates', async (_req, res) => {
  try {
    const data = await fs.readFile(DB_FILE_PATH, 'utf-8');
    const templates = JSON.parse(data);
    // Sort by createdAt descending
    templates.sort((a: any, b: any) => b.createdAt - a.createdAt);
    res.json(templates);
  } catch (error) {
    console.error('Gagal membaca templates dari disk:', error);
    res.status(500).json({ error: 'Gagal membaca database server' });
  }
});

// POST /api/templates
app.post('/api/templates', async (req, res) => {
  try {
    const record = req.body;
    if (!record || !record.id) {
      return res.status(400).json({ error: 'Data template tidak valid' });
    }
    const data = await fs.readFile(DB_FILE_PATH, 'utf-8');
    const templates = JSON.parse(data);
    
    const existingIndex = templates.findIndex((t: any) => t.id === record.id);
    if (existingIndex > -1) {
      templates[existingIndex] = record;
    } else {
      templates.push(record);
    }
    
    await fs.writeFile(DB_FILE_PATH, JSON.stringify(templates, null, 2), 'utf-8');
    res.json({ success: true, record });
  } catch (error) {
    console.error('Gagal menyimpan template ke disk:', error);
    res.status(500).json({ error: 'Gagal menyimpan ke database server' });
  }
});

// DELETE /api/templates/:id
app.delete('/api/templates/:id', async (req, res) => {
  try {
    const id = req.params.id;
    const data = await fs.readFile(DB_FILE_PATH, 'utf-8');
    let templates = JSON.parse(data);
    
    templates = templates.filter((t: any) => t.id !== id);
    
    await fs.writeFile(DB_FILE_PATH, JSON.stringify(templates, null, 2), 'utf-8');
    res.json({ success: true });
  } catch (error) {
    console.error('Gagal menghapus template dari disk:', error);
    res.status(500).json({ error: 'Gagal menghapus dari database server' });
  }
});

// POST /api/convert-to-pdf
app.post('/api/convert-to-pdf', async (req, res) => {
  try {
    const { docxBase64 } = req.body;
    if (!docxBase64) {
      return res.status(400).json({ error: 'Data DOCX base64 tidak boleh kosong' });
    }

    const docxBuf = Buffer.from(docxBase64, 'base64');
    
    // Gunakan libreoffice-convert untuk melakukan konversi
    libre.convert(docxBuf, '.pdf', undefined, (err: any, pdfBuf: any) => {
      if (err) {
        console.error('Gagal konversi dengan libreoffice-convert:', err);
        return res.status(500).json({
          error: 'Gagal mengonversi DOCX ke PDF via LibreOffice Server.',
          details: err.message || String(err),
          code: 'LIBREOFFICE_ERROR'
        });
      }
      
      const pdfBase64 = pdfBuf.toString('base64');
      res.json({ success: true, pdfBase64 });
    });
  } catch (error: any) {
    console.error('Server error saat konversi PDF:', error);
    res.status(500).json({ 
      error: 'Terjadi kesalahan sistem di server saat memproses PDF',
      details: error.message || String(error)
    });
  }
});

// Serve frontend with Vite integration
async function start() {
  await initDB();
  
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }
  
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`[DocuMerge] Server berjalan di http://localhost:${PORT}`);
  });
}

start();
