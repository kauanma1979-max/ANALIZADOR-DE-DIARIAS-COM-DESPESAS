import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';
import 'dotenv/config';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  console.log('Starting server in', process.env.NODE_ENV || 'development', 'mode');

  // Proxy endpoint to bypass CORS for Google Sheets
  app.get('/api/sync-expenses', async (req, res) => {
    try {
      console.log(`[${new Date().toISOString()}] Sync request: ${req.url}`);
      const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lyXkSmeiyyODZbng6GtXTSwNR-XY2KWRLKWqppLef1k/export?format=csv&gid=980751451';
      
      const response = await fetch(SHEET_URL, {
        method: 'GET',
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'text/csv,text/plain,application/csv'
        }
      });
      
      if (!response.ok) {
        const status = response.status;
        const statusText = response.statusText;
        console.error(`Google Sheets Error: ${status} ${statusText}`);
        return res.status(status).json({ 
          error: `Google Sheets responded with ${status}`,
          details: statusText
        });
      }
      
      const csvText = await response.text();
      
      if (!csvText || csvText.length < 10) {
        console.error('Received empty or too short CSV data');
        return res.status(500).json({ error: 'Received empty response from Google Sheets' });
      }

      console.log('Successfully fetched CSV data, length:', csvText.length);
      res.header('Content-Type', 'text/csv; charset=utf-8');
      res.header('Access-Control-Allow-Origin', '*');
      res.send(csvText);
    } catch (error) {
      console.error('Proxy Exception:', error);
      res.status(500).json({ 
        error: 'Failed to fetch spreadsheet data', 
        message: error instanceof Error ? error.message : String(error) 
      });
    }
  });

  // Health check
  app.get('/api/health', (req, res) => {
    res.json({ 
      status: 'ok', 
      time: new Date().toISOString(),
      env: process.env.NODE_ENV
    });
  });

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    // Ensure dist exists or handle gracefully
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'), (err) => {
        if (err) {
          res.status(404).send('Application build not found. Please run build first.');
        }
      });
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server listening at http://0.0.0.0:${PORT}`);
  });
}

startServer().catch(err => {
  console.error('Failed to start server:', err);
});
