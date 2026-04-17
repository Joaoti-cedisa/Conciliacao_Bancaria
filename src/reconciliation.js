/**
 * reconciliation.js — Multi-level reconciliation engine.
 * 
 * Level 1: Group by supplier name, compare sums (exact match with tolerance)
 * Level 2: N:1 / 1:N subset-sum matching for remaining entries
 * Level 3: Value-based cross-supplier matching (same value, different names) → PENDENTE for manual approval
 * Level 4: Mark remaining as unmatched
 */

const TOLERANCE = 0.015; // Allow R$ 0.015 tolerance for rounding

/**
 * Group entries by normalized supplier name.
 */
function groupBySupplier(entries) {
  const groups = {};
  for (const entry of entries) {
    const key = entry.supplierNormalized;
    if (!groups[key]) groups[key] = [];
    groups[key].push(entry);
  }
  return groups;
}

/**
 * Sum values in an array of entries.
 */
function sumValues(entries) {
  return entries.reduce((acc, e) => acc + e.value, 0);
}

/**
 * Round to 2 decimal places.
 */
function round2(n) {
  return Math.round(n * 100) / 100;
}

/**
 * Try to find subsets of `items` that sum to `target` (within tolerance).
 */
function findSubsetSum(items, target, maxIterations = 50000) {
  if (items.length === 0) return null;
  if (items.length > 20) return null;

  let iterations = 0;

  function backtrack(idx, remaining, current) {
    iterations++;
    if (iterations > maxIterations) return null;
    if (Math.abs(remaining) <= TOLERANCE) return [...current];
    // Early exit only if no negative numbers could increase the remaining target
    const hasNegativeRemaining = items.slice(idx).some(i => i.value < 0);
    if (!hasNegativeRemaining && remaining < -TOLERANCE) return null;
    if (idx >= items.length) return null;

    current.push(items[idx]);
    const result = backtrack(idx + 1, round2(remaining - items[idx].value), current);
    if (result) return result;
    current.pop();

    return backtrack(idx + 1, remaining, current);
  }

  const sorted = [...items].sort((a, b) => b.value - a.value);
  return backtrack(0, target, []);
}

/**
 * Run the full reconciliation process.
 * 
 * @param {Array} bankData - Bank entries (already with De-Para applied)
 * @param {Array} fusionData - Fusion entries
 * @returns {Object} { results, valueSuggestions }
 *   - results: Array of reconciliation result objects per supplier
 *   - valueSuggestions: Array of cross-supplier value matches for manual approval
 */
export function reconcile(bankData, fusionData, ignoredSuggestions = []) {
  const bankGroups = groupBySupplier(bankData);
  const fusionGroups = groupBySupplier(fusionData);
  const results = [];

  // Track which suppliers have been matched by name
  const matchedSuppliers = new Set();

  // Get all unique supplier names from both sources
  const allSuppliers = new Set([
    ...Object.keys(bankGroups),
    ...Object.keys(fusionGroups)
  ]);

  for (const supplier of allSuppliers) {
    const bankEntries = bankGroups[supplier] || [];
    const fusionEntries = fusionGroups[supplier] || [];
    const bankTotal = round2(sumValues(bankEntries));
    const fusionTotal = round2(sumValues(fusionEntries));
    const diff = round2(bankTotal - fusionTotal);

    const supplierName = bankEntries[0]?.supplier || fusionEntries[0]?.supplier || supplier;

    // If supplier only exists in one side, skip for now — will handle in value-based matching
    if (bankEntries.length === 0 || fusionEntries.length === 0) {
      continue;
    }

    matchedSuppliers.add(supplier);

    // ===== LEVEL 1: Group Sum Match =====
    if (Math.abs(diff) <= TOLERANCE) {
      results.push({
        id: `r-${results.length}`,
        supplier: supplierName,
        supplierNormalized: supplier,
        bankTotal,
        fusionTotal,
        difference: 0,
        status: 'conciliado',
        level: 1,
        bankEntries,
        fusionEntries,
        matchDetails: null,
        userAction: null // null = auto, 'approved', 'rejected'
      });
      continue;
    }

    // ===== LEVEL 2: N:1 / 1:N Subset Matching =====
    const matchDetails = attemptSubsetMatching(bankEntries, fusionEntries);

    if (matchDetails.unmatchedBank.length === 0 && matchDetails.unmatchedFusion.length === 0) {
      results.push({
        id: `r-${results.length}`,
        supplier: supplierName,
        supplierNormalized: supplier,
        bankTotal,
        fusionTotal,
        difference: 0,
        status: 'conciliado',
        level: 2,
        bankEntries,
        fusionEntries,
        matchDetails,
        userAction: null
      });
      continue;
    }

    // ===== Remaining Differences =====
    results.push({
      id: `r-${results.length}`,
      supplier: supplierName,
      supplierNormalized: supplier,
      bankTotal,
      fusionTotal,
      difference: diff,
      status: 'pendente',
      level: 3,
      bankEntries,
      fusionEntries,
      matchDetails,
      partiallyMatched: matchDetails.matches.length > 0,
      userAction: null
    });
  }

  // ===== LEVEL 3: Value-Based Cross-Supplier Matching =====
  // Collect suppliers that appear only in bank or only in fusion
  const bankOnlySuppliers = Object.keys(bankGroups).filter(s => !matchedSuppliers.has(s));
  const fusionOnlySuppliers = Object.keys(fusionGroups).filter(s => !matchedSuppliers.has(s));

  // Also collect unmatched fusion entries from PENDENTE groups (partial matches from Level 2).
  // Example: ARCELORMITTAL has 19 bank entries and 20 fusion entries — after Level 2, one fusion
  // entry at R$ 111.346,75 remains unpaired because it actually belongs to ALPE FUNDO (bank-only).
  const partialUnmatchedFusionBySupplier = new Map(); // fusionSupplierNormalized -> unmatchedEntries[]
  for (const result of results) {
    if (result.status === 'pendente' && result.matchDetails && result.matchDetails.unmatchedFusion.length > 0) {
      partialUnmatchedFusionBySupplier.set(result.supplierNormalized, result.matchDetails.unmatchedFusion);
    }
  }

  const valueSuggestions = [];
  const usedBankSuppliers = new Set();
  const usedFusionSuppliers = new Set();
  // Tracks which partial (PENDENTE) fusion entries are already claimed by a suggestion.
  const usedPartialFusionEntries = new Set(); // uses entry object identity

  // --- Pass A: match by total value against fusion-only suppliers ---
  for (const bankKey of bankOnlySuppliers) {
    if (usedBankSuppliers.has(bankKey)) continue;
    const bankEntries = bankGroups[bankKey];
    const bankTotal = round2(sumValues(bankEntries));

    for (const fusionKey of fusionOnlySuppliers) {
      if (usedFusionSuppliers.has(fusionKey)) continue;

      if (ignoredSuggestions.some(ig => ig.nome_banco_normalizado === bankKey && ig.nome_fusion_normalizado === fusionKey)) {
        continue;
      }

      const fusionEntries = fusionGroups[fusionKey];
      const fusionTotal = round2(sumValues(fusionEntries));

      if (Math.abs(bankTotal - fusionTotal) <= TOLERANCE) {
        valueSuggestions.push({
          id: `vs-${valueSuggestions.length}`,
          bankSupplier: bankEntries[0]?.supplier || bankKey,
          bankSupplierNormalized: bankKey,
          fusionSupplier: fusionEntries[0]?.supplier || fusionKey,
          fusionSupplierNormalized: fusionKey,
          bankTotal,
          fusionTotal,
          difference: 0,
          bankEntries,
          fusionEntries,
          matchType: 'value_total',
          status: 'sugerido',
          userAction: null
        });
        usedBankSuppliers.add(bankKey);
        usedFusionSuppliers.add(fusionKey);
        break;
      }
    }
  }

  // --- Pass B: match by individual entry values against fusion-only suppliers ---
  for (const bankKey of bankOnlySuppliers) {
    if (usedBankSuppliers.has(bankKey)) continue;
    const bankEntries = bankGroups[bankKey];

    for (const fusionKey of fusionOnlySuppliers) {
      if (usedFusionSuppliers.has(fusionKey)) continue;

      if (ignoredSuggestions.some(ig => ig.nome_banco_normalizado === bankKey && ig.nome_fusion_normalizado === fusionKey)) {
        continue;
      }

      const fusionEntries = fusionGroups[fusionKey];

      const valueMatches = [];
      for (const be of bankEntries) {
        for (const fe of fusionEntries) {
          if (Math.abs(be.value - fe.value) <= TOLERANCE) {
            valueMatches.push({ bank: be, fusion: fe });
          }
        }
      }

      if (valueMatches.length > 0) {
        const bankTotal = round2(sumValues(bankEntries));
        const fusionTotal = round2(sumValues(fusionEntries));
        valueSuggestions.push({
          id: `vs-${valueSuggestions.length}`,
          bankSupplier: bankEntries[0]?.supplier || bankKey,
          bankSupplierNormalized: bankKey,
          fusionSupplier: fusionEntries[0]?.supplier || fusionKey,
          fusionSupplierNormalized: fusionKey,
          bankTotal,
          fusionTotal,
          difference: round2(bankTotal - fusionTotal),
          bankEntries,
          fusionEntries,
          valueMatches,
          matchType: 'value_line',
          status: 'sugerido',
          userAction: null
        });
        usedBankSuppliers.add(bankKey);
        usedFusionSuppliers.add(fusionKey);
        break;
      }
    }
  }

  // --- Pass C: match bank-only entries against UNMATCHED fusion entries inside PENDENTE groups ---
  // This handles the case where, e.g., ALPE FUNDO (bank-only) has the same value as a fusion entry
  // that lives inside the ARCELORMITTAL PENDENTE group but was never paired during Level 2.
  for (const bankKey of bankOnlySuppliers) {
    if (usedBankSuppliers.has(bankKey)) continue;
    const bankEntries = bankGroups[bankKey];
    const bankTotal = round2(sumValues(bankEntries));

    for (const [fusionKey, unmatchedFusionEntries] of partialUnmatchedFusionBySupplier) {
      if (ignoredSuggestions.some(ig => ig.nome_banco_normalizado === bankKey && ig.nome_fusion_normalizado === fusionKey)) {
        continue;
      }

      // Only consider orphan entries not already claimed by a prior suggestion.
      const available = unmatchedFusionEntries.filter(e => !usedPartialFusionEntries.has(e));
      if (available.length === 0) continue;

      const availableTotal = round2(sumValues(available));

      // Sub-pass C1: bank total matches the sum of ALL available orphan fusion entries.
      if (Math.abs(bankTotal - availableTotal) <= TOLERANCE) {
        available.forEach(e => usedPartialFusionEntries.add(e));
        valueSuggestions.push({
          id: `vs-${valueSuggestions.length}`,
          bankSupplier: bankEntries[0]?.supplier || bankKey,
          bankSupplierNormalized: bankKey,
          fusionSupplier: fusionGroups[fusionKey]?.[0]?.supplier || fusionKey,
          fusionSupplierNormalized: fusionKey,
          bankTotal,
          fusionTotal: availableTotal,
          difference: 0,
          bankEntries,
          fusionEntries: available,
          matchType: 'value_total_partial',
          status: 'sugerido',
          userAction: null,
          isPartialFusionGroup: true
        });
        usedBankSuppliers.add(bankKey);
        break;
      }

      // Sub-pass C2: find a subset of available orphan entries that sums to bankTotal.
      if (available.length <= 20) {
        const subsetMatch = findSubsetSum(available, bankTotal);
        if (subsetMatch) {
          subsetMatch.forEach(e => usedPartialFusionEntries.add(e));
          const subsetTotal = round2(sumValues(subsetMatch));
          valueSuggestions.push({
            id: `vs-${valueSuggestions.length}`,
            bankSupplier: bankEntries[0]?.supplier || bankKey,
            bankSupplierNormalized: bankKey,
            fusionSupplier: fusionGroups[fusionKey]?.[0]?.supplier || fusionKey,
            fusionSupplierNormalized: fusionKey,
            bankTotal,
            fusionTotal: subsetTotal,
            difference: 0,
            bankEntries,
            fusionEntries: subsetMatch,
            matchType: 'value_total_partial',
            status: 'sugerido',
            userAction: null,
            isPartialFusionGroup: true
          });
          usedBankSuppliers.add(bankKey);
          break;
        }
      }

      // Sub-pass C3: any individual bank entry matches any available orphan fusion entry by value.
      const valueMatches = [];
      const usedFusionInMatch = new Set();
      for (const be of bankEntries) {
        for (const fe of available) {
          if (!usedFusionInMatch.has(fe) && Math.abs(be.value - fe.value) <= TOLERANCE) {
            valueMatches.push({ bank: be, fusion: fe });
            usedFusionInMatch.add(fe);
            break;
          }
        }
      }

      if (valueMatches.length > 0) {
        valueMatches.forEach(m => usedPartialFusionEntries.add(m.fusion));
        const matchedFusionEntries = valueMatches.map(m => m.fusion);
        const matchedFusionTotal = round2(sumValues(matchedFusionEntries));
        valueSuggestions.push({
          id: `vs-${valueSuggestions.length}`,
          bankSupplier: bankEntries[0]?.supplier || bankKey,
          bankSupplierNormalized: bankKey,
          fusionSupplier: fusionGroups[fusionKey]?.[0]?.supplier || fusionKey,
          fusionSupplierNormalized: fusionKey,
          bankTotal,
          fusionTotal: matchedFusionTotal,
          difference: round2(bankTotal - matchedFusionTotal),
          bankEntries,
          fusionEntries: matchedFusionEntries,
          valueMatches,
          matchType: 'value_line_partial',
          status: 'sugerido',
          userAction: null,
          isPartialFusionGroup: true
        });
        usedBankSuppliers.add(bankKey);
        break;
      }
    }
  }

  // Remaining unmatched suppliers (no name or value match found)
  for (const bankKey of bankOnlySuppliers) {
    if (usedBankSuppliers.has(bankKey)) continue;
    const bankEntries = bankGroups[bankKey];
    const bankTotal = round2(sumValues(bankEntries));

    results.push({
      id: `r-${results.length}`,
      supplier: bankEntries[0]?.supplier || bankKey,
      supplierNormalized: bankKey,
      bankTotal,
      fusionTotal: 0,
      difference: bankTotal,
      status: 'pendente',
      level: 4,
      bankEntries,
      fusionEntries: [],
      matchDetails: null,
      onlyInBank: true,
      userAction: null
    });
  }

  for (const fusionKey of fusionOnlySuppliers) {
    if (usedFusionSuppliers.has(fusionKey)) continue;
    const fusionEntries = fusionGroups[fusionKey];
    const fusionTotal = round2(sumValues(fusionEntries));

    results.push({
      id: `r-${results.length}`,
      supplier: fusionEntries[0]?.supplier || fusionKey,
      supplierNormalized: fusionKey,
      bankTotal: 0,
      fusionTotal,
      difference: -fusionTotal,
      status: 'pendente',
      level: 4,
      bankEntries: [],
      fusionEntries,
      matchDetails: null,
      onlyInFusion: true,
      userAction: null
    });
  }

  // Sort: sugerido first, then pendente, then conciliado
  results.sort((a, b) => {
    const order = { pendente: 0, conciliado: 1 };
    if ((order[a.status] ?? 0) !== (order[b.status] ?? 0)) {
      return (order[a.status] ?? 0) - (order[b.status] ?? 0);
    }
    return Math.abs(b.difference) - Math.abs(a.difference);
  });

  return { results, valueSuggestions };
}

/**
 * Attempt N:1 / 1:N matching between bank and fusion entries.
 */
function attemptSubsetMatching(bankEntries, fusionEntries) {
  const matches = [];
  let unmatchedBank = [...bankEntries];
  let unmatchedFusion = [...fusionEntries];

  // Pass 1: Exact 1:1 matches
  for (let i = unmatchedBank.length - 1; i >= 0; i--) {
    for (let j = unmatchedFusion.length - 1; j >= 0; j--) {
      if (Math.abs(unmatchedBank[i].value - unmatchedFusion[j].value) <= TOLERANCE) {
        matches.push({
          type: '1:1',
          bank: [unmatchedBank[i]],
          fusion: [unmatchedFusion[j]]
        });
        unmatchedBank.splice(i, 1);
        unmatchedFusion.splice(j, 1);
        break;
      }
    }
  }

  // Pass 2: N:1 — multiple bank entries → one fusion entry
  for (let j = unmatchedFusion.length - 1; j >= 0; j--) {
    const target = unmatchedFusion[j].value;
    const subset = findSubsetSum(unmatchedBank, target);
    if (subset && subset.length > 1) {
      matches.push({
        type: 'N:1',
        bank: subset,
        fusion: [unmatchedFusion[j]]
      });
      for (const s of subset) {
        const idx = unmatchedBank.findIndex(e => e === s);
        if (idx !== -1) unmatchedBank.splice(idx, 1);
      }
      unmatchedFusion.splice(j, 1);
    }
  }

  // Pass 3: 1:N — one bank entry → multiple fusion entries
  for (let i = unmatchedBank.length - 1; i >= 0; i--) {
    const target = unmatchedBank[i].value;
    const subset = findSubsetSum(unmatchedFusion, target);
    if (subset && subset.length > 1) {
      matches.push({
        type: '1:N',
        bank: [unmatchedBank[i]],
        fusion: subset
      });
      for (const s of subset) {
        const idx = unmatchedFusion.findIndex(e => e === s);
        if (idx !== -1) unmatchedFusion.splice(idx, 1);
      }
      unmatchedBank.splice(i, 1);
    }
  }

  return { matches, unmatchedBank, unmatchedFusion };
}

/**
 * Compute summary statistics from reconciliation results.
 */
export function computeStats(results, valueSuggestions = []) {
  const conciliado = results.filter(r => r.status === 'conciliado' || r.userAction === 'approved');
  const pendente = results.filter(r => r.status === 'pendente' && r.userAction !== 'approved');
  const sugerido = valueSuggestions.filter(s => s.userAction === null);
  const aprovado = valueSuggestions.filter(s => s.userAction === 'approved');
  const totalDiff = round2(
    pendente.reduce((acc, r) => acc + r.difference, 0) +
    sugerido.reduce((acc, s) => acc + s.difference, 0)
  );

  return {
    total: results.length + valueSuggestions.length,
    conciliado: conciliado.length + aprovado.length,
    pendente: pendente.length,
    sugerido: sugerido.length,
    totalDifference: totalDiff,
    totalBank: round2(results.reduce((acc, r) => acc + r.bankTotal, 0)),
    totalFusion: round2(results.reduce((acc, r) => acc + r.fusionTotal, 0))
  };
}
