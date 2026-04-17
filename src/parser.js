/**
 * parser.js — Excel file parser for Bank and Fusion files.
 * Uses SheetJS (xlsx) to parse and normalize data.
 */
import * as XLSX from 'xlsx';

/**
 * Parse a Brazilian number string to float.
 * Handles formats like: "1.234,56" → 1234.56, "1234.56" → 1234.56
 */
export function parseBrazilianNumber(value) {
  if (value == null) return 0;
  if (typeof value === 'number') return value;

  let str = String(value).trim();
  if (!str) return 0;

  // Remove currency symbols, spaces
  str = str.replace(/[R$\s]/g, '');

  // Detect Brazilian format (dot as thousands, comma as decimal)
  if (/\d{1,3}(\.\d{3})*(,\d{1,2})?$/.test(str)) {
    str = str.replace(/\./g, '').replace(',', '.');
  }
  // Handle comma as decimal without dots: "1234,56"
  else if (/^\d+(,\d{1,2})$/.test(str)) {
    str = str.replace(',', '.');
  }

  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

/**
 * Normalize a supplier name for matching purposes.
 */
export function normalizeName(name) {
  if (!name) return '';
  return String(name)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .replace(/[.\-\/]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Find column index by trying multiple possible header names.
 * Searches across ALL provided rows (handles multi-row headers).
 */
function findColumn(headerRows, possibleNames) {
  for (const name of possibleNames) {
    for (const headers of headerRows) {
      const idx = headers.findIndex(h =>
        String(h || '').toLowerCase().includes(name.toLowerCase())
      );
      if (idx !== -1) return idx;
    }
  }
  return -1;
}

/**
 * Parse the Bank (DDA) Excel file.
 * 
 * Known structure (Banco Safra):
 * - Rows 0-8: Bank header info (agency, account, etc.)
 * - Row 9: Column headers
 * - Row 10+: Data
 * 
 * Returns: [{ supplier, value, date, document, modality, raw }]
 */
export function parseBankFile(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // Find the header row — look for a row containing "Favorecido" or "Beneficiário"
  let headerIdx = -1;
  for (let i = 0; i < Math.min(data.length, 20); i++) {
    const row = data[i];
    if (!row) continue;
    const rowStr = row.map(c => String(c || '').toLowerCase()).join(' ');
    if (rowStr.includes('favorecido') || rowStr.includes('beneficiário') || rowStr.includes('beneficiario')) {
      headerIdx = i;
      break;
    }
  }

  if (headerIdx === -1) {
    throw new Error('Cabeçalho não encontrado no arquivo do Banco. Procurando por "Favorecido / Beneficiário".');
  }

  const headers = data[headerIdx] || [];

  // Find relevant columns
  const colSupplier = findColumn([headers], ['favorecido', 'beneficiário', 'beneficiario']);
  const colValue = findColumn([headers], ['valor (r$)', 'valor r$', 'valor']);
  const colDate = findColumn([headers], ['data']);
  const colDocument = findColumn([headers], ['n° documento', 'nº documento', 'documento']);
  const colModality = findColumn([headers], ['modalidade']);

  if (colSupplier === -1) {
    throw new Error(`Coluna de fornecedor/beneficiário não encontrada. Cabeçalhos encontrados: ${headers.filter(h => h).join(', ')}`);
  }
  if (colValue === -1) {
    throw new Error(`Coluna de valor não encontrada. Cabeçalhos encontrados: ${headers.filter(h => h).join(', ')}`);
  }

  console.log(`[Bank Parser] Header row: ${headerIdx}, Supplier col: ${colSupplier}, Value col: ${colValue}`);

  const results = [];
  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[colSupplier]) continue;

    let supplier = String(row[colSupplier]).trim();
    const value = parseBrazilianNumber(row[colValue]);

    if (!supplier || value === 0) continue;

    // Clean supplier name: remove leading CNPJ/number prefix if present
    // e.g., "30.401.462 ANDERSON DA SILVA" → "ANDERSON DA SILVA"
    const cnpjPrefixMatch = supplier.match(/^[\d.\/\-]+\s+(.+)$/);
    if (cnpjPrefixMatch) {
      supplier = cnpjPrefixMatch[1];
    }

    results.push({
      supplier,
      supplierNormalized: normalizeName(supplier),
      value: Math.abs(value),
      date: row[colDate] || '',
      document: row[colDocument] || '',
      modality: row[colModality] || '',
      source: 'bank',
      raw: row
    });
  }

  console.log(`[Bank Parser] Parsed ${results.length} entries`);
  return results;
}

/**
 * Parse the Fusion (DIA) Excel file.
 * 
 * Known structure:
 * - Row 0: Group headers ("Fornecedor ou Parte", "NFF", "Pagamento")
 * - Row 1: Sub-headers ("Número", "Data de Vencimento", "Moeda", etc.)
 * - Row 2+: Data
 * - Column A (index 0) is empty; data starts at column B (index 1)
 * 
 * Returns: [{ supplier, value, date, number, raw }]
 */
export function parseFusionFile(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // Find header rows — look for "Fornecedor" in any of the first 15 rows
  let mainHeaderIdx = -1;
  for (let i = 0; i < Math.min(data.length, 15); i++) {
    const row = data[i];
    if (!row) continue;
    const rowStr = row.map(c => String(c || '').toLowerCase()).join(' ');
    if (rowStr.includes('fornecedor')) {
      mainHeaderIdx = i;
      break;
    }
  }

  if (mainHeaderIdx === -1) {
    throw new Error('Cabeçalho não encontrado no arquivo do Fusion. Procurando por "Fornecedor".');
  }

  // Build merged headers from main header row + sub-header row(s)
  const headerRows = [];
  headerRows.push(data[mainHeaderIdx] || []);
  
  // Check if next row(s) are sub-headers (contain things like "Número", "Data de Vencimento", etc.)
  let subHeaderEnd = mainHeaderIdx;
  for (let i = mainHeaderIdx + 1; i < Math.min(mainHeaderIdx + 3, data.length); i++) {
    const row = data[i];
    if (!row) break;
    const rowStr = row.map(c => String(c || '').toLowerCase()).join(' ');
    if (rowStr.includes('número') || rowStr.includes('numero') || rowStr.includes('data de vencimento') || 
        rowStr.includes('moeda') || rowStr.includes('valor')) {
      headerRows.push(row);
      subHeaderEnd = i;
    } else {
      break;
    }
  }

  console.log(`[Fusion Parser] Main header row: ${mainHeaderIdx}, Sub-headers through row: ${subHeaderEnd}`);
  console.log(`[Fusion Parser] Header rows content:`, headerRows.map(r => r.filter(c => c)));

  // Find supplier column in main header
  const colSupplier = findColumn(headerRows, ['fornecedor ou parte', 'fornecedor']);

  // Find value column — prefer "Valor com Juros" (the payment amount)
  let colValue = findColumn(headerRows, ['valor com juros']);
  if (colValue === -1) {
    colValue = findColumn(headerRows, ['valor não pago', 'valor nao pago']);
  }
  if (colValue === -1) {
    // Fallback: Find "Valor" but avoid matching "Valor Não Pago" etc.
    // Look at sub-headers for standalone "Valor"
    for (const row of headerRows) {
      for (let j = 0; j < row.length; j++) {
        const h = String(row[j] || '').trim().toLowerCase();
        if (h === 'valor') {
          colValue = j;
          break;
        }
      }
      if (colValue !== -1) break;
    }
  }

  const colDate = findColumn(headerRows, ['data de vencimento', 'vencimento']);
  const colNumber = findColumn(headerRows, ['número', 'numero']);

  if (colSupplier === -1) {
    const allHeaders = headerRows.flat().filter(h => h);
    throw new Error(`Coluna de fornecedor não encontrada no Fusion. Cabeçalhos: ${allHeaders.join(', ')}`);
  }
  if (colValue === -1) {
    const allHeaders = headerRows.flat().filter(h => h);
    throw new Error(`Coluna de valor não encontrada no Fusion. Cabeçalhos: ${allHeaders.join(', ')}`);
  }

  console.log(`[Fusion Parser] Supplier col: ${colSupplier}, Value col: ${colValue}, Date col: ${colDate}, Number col: ${colNumber}`);

  // Data starts after all header rows
  const dataStart = subHeaderEnd + 1;

  const results = [];
  for (let i = dataStart; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[colSupplier]) continue;

    const supplierRaw = String(row[colSupplier]).trim();
    // Skip empty rows or rows that look like subtotals
    if (!supplierRaw || supplierRaw.toLowerCase().includes('total')) continue;

    // Clean supplier name: remove leading CNPJ/number prefix if present
    // e.g., "60.852.801 FRANSOAR BARBOZA MARTINS" → "FRANSOAR BARBOZA MARTINS"
    let supplier = supplierRaw;
    const cnpjPrefixMatch = supplier.match(/^[\d.\/\-]+\s+(.+)$/);
    if (cnpjPrefixMatch) {
      supplier = cnpjPrefixMatch[1];
    }

    const value = parseBrazilianNumber(row[colValue]);
    if (value === 0) continue;

    results.push({
      supplier,
      supplierNormalized: normalizeName(supplier),
      value: value,
      date: row[colDate] || '',
      number: row[colNumber] || '',
      source: 'fusion',
      raw: row
    });
  }

  console.log(`[Fusion Parser] Parsed ${results.length} entries`);
  return results;
}
