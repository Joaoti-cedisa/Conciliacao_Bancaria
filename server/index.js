/**
 * server/index.js — Express backend with SQLite for De-Para persistence.
 */
const express = require('express');
const Database = require('better-sqlite3');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = 3001;

// Middleware
app.use(cors());
app.use(express.json());

// ===== SQLite Database Setup =====
const dbPath = path.join(__dirname, 'conciliacao.db');
const db = new Database(dbPath);

// Enable WAL mode for better performance
db.pragma('journal_mode = WAL');

// Create De-Para table
db.exec(`
  CREATE TABLE IF NOT EXISTS depara (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_banco TEXT NOT NULL,
    nome_banco_normalizado TEXT NOT NULL UNIQUE,
    nome_fusion TEXT NOT NULL,
    criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
    atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
  )
`);

db.exec(`
  CREATE TABLE IF NOT EXISTS sugestoes_recusadas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_banco_normalizado TEXT NOT NULL,
    nome_fusion_normalizado TEXT NOT NULL,
    criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(nome_banco_normalizado, nome_fusion_normalizado)
  )
`);

console.log('✅ Database initialized at:', dbPath);

// ===== API Endpoints =====

/**
 * GET /api/depara — List all De-Para mappings
 */
app.get('/api/depara', (req, res) => {
  try {
    const rows = db.prepare('SELECT * FROM depara ORDER BY nome_banco ASC').all();
    res.json({ success: true, data: rows });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * POST /api/depara — Create a new De-Para mapping
 * Body: { nome_banco, nome_banco_normalizado, nome_fusion }
 */
app.post('/api/depara', (req, res) => {
  try {
    const { nome_banco, nome_banco_normalizado, nome_fusion } = req.body;

    if (!nome_banco || !nome_banco_normalizado || !nome_fusion) {
      return res.status(400).json({
        success: false,
        error: 'Campos obrigatórios: nome_banco, nome_banco_normalizado, nome_fusion'
      });
    }

    const stmt = db.prepare(`
      INSERT INTO depara (nome_banco, nome_banco_normalizado, nome_fusion)
      VALUES (?, ?, ?)
      ON CONFLICT(nome_banco_normalizado) DO UPDATE SET
        nome_fusion = excluded.nome_fusion,
        atualizado_em = CURRENT_TIMESTAMP
    `);

    const result = stmt.run(nome_banco, nome_banco_normalizado, nome_fusion);
    res.json({ success: true, id: result.lastInsertRowid });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * POST /api/depara/batch — Create multiple De-Para mappings at once
 * Body: [{ nome_banco, nome_banco_normalizado, nome_fusion }]
 */
app.post('/api/depara/batch', (req, res) => {
  try {
    const mappings = req.body;

    if (!Array.isArray(mappings) || mappings.length === 0) {
      return res.status(400).json({
        success: false,
        error: 'Body deve ser um array de mapeamentos'
      });
    }

    const stmt = db.prepare(`
      INSERT INTO depara (nome_banco, nome_banco_normalizado, nome_fusion)
      VALUES (?, ?, ?)
      ON CONFLICT(nome_banco_normalizado) DO UPDATE SET
        nome_fusion = excluded.nome_fusion,
        atualizado_em = CURRENT_TIMESTAMP
    `);

    const insertMany = db.transaction((items) => {
      for (const item of items) {
        stmt.run(item.nome_banco, item.nome_banco_normalizado, item.nome_fusion);
      }
    });

    insertMany(mappings);
    res.json({ success: true, count: mappings.length });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * PUT /api/depara/:id — Update a De-Para mapping
 * Body: { nome_fusion }
 */
app.put('/api/depara/:id', (req, res) => {
  try {
    const { id } = req.params;
    const { nome_fusion } = req.body;

    if (!nome_fusion) {
      return res.status(400).json({
        success: false,
        error: 'Campo obrigatório: nome_fusion'
      });
    }

    const stmt = db.prepare(`
      UPDATE depara SET nome_fusion = ?, atualizado_em = CURRENT_TIMESTAMP
      WHERE id = ?
    `);

    const result = stmt.run(nome_fusion, id);
    if (result.changes === 0) {
      return res.status(404).json({ success: false, error: 'Mapeamento não encontrado' });
    }

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * DELETE /api/depara/:id — Delete a De-Para mapping
 */
app.delete('/api/depara/:id', (req, res) => {
  try {
    const { id } = req.params;
    const result = db.prepare('DELETE FROM depara WHERE id = ?').run(id);

    if (result.changes === 0) {
      return res.status(404).json({ success: false, error: 'Mapeamento não encontrado' });
    }

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * GET /api/depara/search?q=term — Search mappings
 */
app.get('/api/depara/search', (req, res) => {
  try {
    const { q } = req.query;
    if (!q) {
      return res.json({ success: true, data: [] });
    }

    const rows = db.prepare(`
      SELECT * FROM depara
      WHERE nome_banco LIKE ? OR nome_fusion LIKE ?
      ORDER BY nome_banco ASC
    `).all(`%${q}%`, `%${q}%`);

    res.json({ success: true, data: rows });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * GET /api/depara/dictionary — Get all mappings as a simple dictionary { normalizedName: fusionName }
 */
app.get('/api/depara/dictionary', (req, res) => {
  try {
    const rows = db.prepare('SELECT nome_banco_normalizado, nome_fusion FROM depara').all();
    const dict = {};
    for (const row of rows) {
      dict[row.nome_banco_normalizado] = row.nome_fusion;
    }
    res.json({ success: true, data: dict });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ===== API Endpoints: Sugestões Recusadas =====

/**
 * GET /api/depara/recusadas - List all ignored suggestions
 */
app.get('/api/depara/recusadas', (req, res) => {
  try {
    const rows = db.prepare('SELECT * FROM sugestoes_recusadas ORDER BY criado_em DESC').all();
    res.json({ success: true, data: rows });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * POST /api/depara/recusadas - Add an ignored suggestion
 */
app.post('/api/depara/recusadas', (req, res) => {
  try {
    const { nome_banco_normalizado, nome_fusion_normalizado } = req.body;
    if (!nome_banco_normalizado || !nome_fusion_normalizado) {
      return res.status(400).json({ success: false, error: 'Campos obrigatórios: nome_banco_normalizado, nome_fusion_normalizado' });
    }

    const stmt = db.prepare(`
      INSERT OR IGNORE INTO sugestoes_recusadas (nome_banco_normalizado, nome_fusion_normalizado)
      VALUES (?, ?)
    `);
    const result = stmt.run(nome_banco_normalizado, nome_fusion_normalizado);
    res.json({ success: true, id: result.lastInsertRowid });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * DELETE /api/depara/recusadas/:id - Delete an ignored suggestion
 */
app.delete('/api/depara/recusadas/:id', (req, res) => {
  try {
    const { id } = req.params;
    const result = db.prepare('DELETE FROM sugestoes_recusadas WHERE id = ?').run(id);
    if (result.changes === 0) {
      return res.status(404).json({ success: false, error: 'Recusa não encontrada' });
    }
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ===== Start Server =====

app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
  console.log(`📦 Database: ${dbPath}`);
});

// Graceful shutdown
process.on('SIGINT', () => {
  db.close();
  process.exit(0);
});
