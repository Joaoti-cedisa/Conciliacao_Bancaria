/**
 * depara.js — Supplier name mapping (De-Para) dictionary.
 * Now uses SQLite backend via REST API for persistent storage.
 */

const API_BASE = '/api/depara';
import { normalizeName } from './parser.js';

/**
 * Load all De-Para mappings from the database.
 * Returns: { normalizedBankName: fusionName, ... }
 */
export async function loadDictionary() {
  try {
    const res = await fetch(`${API_BASE}/dictionary`, { cache: 'no-store' });
    const json = await res.json();
    return json.success ? json.data : {};
  } catch {
    console.error('Failed to load De-Para dictionary from server');
    return {};
  }
}

/**
 * Load all De-Para records (full rows with id, names, timestamps).
 */
export async function loadAllRecords() {
  try {
    const res = await fetch(API_BASE, { cache: 'no-store' });
    const json = await res.json();
    return json.success ? json.data : [];
  } catch {
    console.error('Failed to load De-Para records from server');
    return [];
  }
}

/**
 * Save a single mapping to the database.
 */
export async function setMapping(bankName, bankNameNormalized, fusionName) {
  try {
    const res = await fetch(API_BASE, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        nome_banco: bankName,
        nome_banco_normalizado: bankNameNormalized,
        nome_fusion: fusionName
      })
    });
    const json = await res.json();
    return json.success;
  } catch (err) {
    console.error('Failed to save mapping:', err);
    return false;
  }
}

/**
 * Save multiple mappings at once.
 */
export async function setMappingBatch(mappings) {
  try {
    const res = await fetch(`${API_BASE}/batch`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(mappings.map(m => ({
        nome_banco: m.bankName,
        nome_banco_normalizado: m.bankNormalized,
        nome_fusion: m.fusionName
      })))
    });
    const json = await res.json();
    return json.success;
  } catch (err) {
    console.error('Failed to save batch mappings:', err);
    return false;
  }
}

/**
 * Update an existing mapping by its ID.
 */
export async function updateMapping(id, fusionName) {
  try {
    const res = await fetch(`${API_BASE}/${id}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ nome_fusion: fusionName })
    });
    const json = await res.json();
    return json.success;
  } catch (err) {
    console.error('Failed to update mapping:', err);
    return false;
  }
}

/**
 * Delete a mapping by its ID.
 */
export async function deleteMapping(id) {
  try {
    const res = await fetch(`${API_BASE}/${id}`, { method: 'DELETE' });
    const json = await res.json();
    return json.success;
  } catch (err) {
    console.error('Failed to delete mapping:', err);
    return false;
  }
}

/**
 * Search mappings.
 */
export async function searchMappings(query) {
  try {
    const res = await fetch(`${API_BASE}/search?q=${encodeURIComponent(query)}`, { cache: 'no-store' });
    const json = await res.json();
    return json.success ? json.data : [];
  } catch {
    return [];
  }
}

/**
 * Compute Levenshtein distance between two strings.
 */
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;

  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }
  return dp[m][n];
}

/**
 * Find the best fuzzy match for a bank name among Fusion names.
 */
export function findBestMatch(bankName, fusionNames) {
  let bestMatch = null;
  let bestScore = Infinity;
  const bankUpper = bankName.toUpperCase();

  for (const fn of fusionNames) {
    const fusionUpper = fn.toUpperCase();

    if (bankUpper === fusionUpper) {
      return { name: fn, score: 0 };
    }

    if (bankUpper.includes(fusionUpper) || fusionUpper.includes(bankUpper)) {
      const score = 1;
      if (score < bestScore) {
        bestScore = score;
        bestMatch = { name: fn, score };
      }
      continue;
    }

    const dist = levenshtein(bankUpper, fusionUpper);
    const maxLen = Math.max(bankUpper.length, fusionUpper.length);
    const similarity = 1 - dist / maxLen;

    if (similarity >= 0.6 && dist < bestScore) {
      bestScore = dist;
      bestMatch = { name: fn, score: dist };
    }
  }

  return bestMatch;
}

/**
 * Analyze mappings between bank and fusion data.
 * Returns: { unmapped, mapped, autoMatched }
 */
export async function analyzeMapping(bankData, fusionData) {
  const dictionary = await loadDictionary();

  const bankNames = [...new Set(bankData.map(d => d.supplierNormalized))];
  const fusionNames = [...new Set(fusionData.map(d => d.supplier))];
  const fusionNormalized = new Set(fusionData.map(d => d.supplierNormalized));

  const unmapped = [];
  const mapped = [];
  const autoMatched = [];
  const autoSaveBatch = [];

  for (const bankNorm of bankNames) {
    const originalName = bankData.find(d => d.supplierNormalized === bankNorm)?.supplier || bankNorm;

    // Check if already in dictionary
    if (dictionary[bankNorm]) {
      mapped.push({
        bankName: originalName,
        bankNormalized: bankNorm,
        fusionName: dictionary[bankNorm],
        source: 'dictionary'
      });
      continue;
    }

    // Check if exact normalized match exists in Fusion
    if (fusionNormalized.has(bankNorm)) {
      const fusionOriginal = fusionData.find(d => d.supplierNormalized === bankNorm)?.supplier || bankNorm;
      autoMatched.push({
        bankName: originalName,
        bankNormalized: bankNorm,
        fusionName: fusionOriginal,
        source: 'auto'
      });
      autoSaveBatch.push({
        bankName: originalName,
        bankNormalized: bankNorm,
        fusionName: fusionOriginal
      });
      continue;
    }

    // Try fuzzy matching
    const suggestion = findBestMatch(bankNorm, fusionNames);
    unmapped.push({
      bankName: originalName,
      bankNormalized: bankNorm,
      suggestion: suggestion ? suggestion.name : null
    });
  }

  // Auto-save exact matches to DB
  if (autoSaveBatch.length > 0) {
    await setMappingBatch(autoSaveBatch);
  }

  return { unmapped, mapped, autoMatched };
}

/**
 * Apply the De-Para dictionary to bank data entries.
 */
export async function applyDictionaryToData(bankData) {
  const dictionary = await loadDictionary();

  return bankData.map(entry => {
    const mappedName = dictionary[entry.supplierNormalized];
    if (mappedName) {
      return {
        ...entry,
        supplierOriginal: entry.supplier,
        supplier: mappedName,
        supplierNormalized: normalizeName(mappedName)
      };
    }
    return { ...entry, supplierOriginal: entry.supplier };
  });
}

/**
 * Load all ignored suggestions (rejected by user).
 */
export async function loadIgnoredSuggestions() {
  try {
    const res = await fetch(`${API_BASE}/recusadas`, { cache: 'no-store' });
    const json = await res.json();
    return json.success ? json.data : [];
  } catch {
    console.error('Failed to load ignored suggestions');
    return [];
  }
}

/**
 * Ignore a value suggestion (reject).
 */
export async function ignoreSuggestion(bankNorm, fusionNorm) {
  try {
    const res = await fetch(`${API_BASE}/recusadas`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        nome_banco_normalizado: bankNorm,
        nome_fusion_normalizado: fusionNorm
      })
    });
    const json = await res.json();
    return json.success;
  } catch (err) {
    console.error('Failed to save ignored suggestion:', err);
    return false;
  }
}

/**
 * Remove an ignored suggestion by its ID (undo).
 */
export async function unignoreSuggestion(id) {
  try {
    const res = await fetch(`${API_BASE}/recusadas/${id}`, { method: 'DELETE' });
    const json = await res.json();
    return json.success;
  } catch (err) {
    console.error('Failed to remove ignored suggestion:', err);
    return false;
  }
}
