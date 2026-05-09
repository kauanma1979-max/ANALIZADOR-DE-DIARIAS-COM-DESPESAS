import express from 'express';
import { createServer as createViteServer } from 'vite';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Proxy endpoint to bypass CORS for Google Sheets
  app.get('/api/sync-expenses', async (req, res) => {
    try {
      console.log('Received sync request at /api/sync-expenses');
      const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1lyXkSmeiyyODZbng6GtXTSwNR-XY2KWRLKWqppLef1k/export?format=csv&gid=980751451';
      
      const response = await fetch(SHEET_URL, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }
      });
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Google Sheets Error: ${response.status} - ${errorText}`);
        return res.status(response.status).json({ 
          error: `Google Sheets responded with ${response.status}`,
          details: errorText.substring(0, 500)
        });
      }
      
      const csvText = await response.text();
      console.log('Successfully fetched CSV data, length:', csvText.length);
      res.header('Content-Type', 'text/csv');
      res.send(csvText);
    } catch (error) {
      console.error('Proxy Error:', error);
      res.status(500).json({ error: 'Failed to fetch spreadsheet data', message: error instanceof Error ? error.message : String(error) });
    }
  });

  // Health check
  app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', time: new Date().toISOString() });
  });

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
