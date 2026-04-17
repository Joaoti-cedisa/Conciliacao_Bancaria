/**
 * main.js — Application entry point.
 * Wires up UI events, file uploads, De-Para sidebar, and reconciliation flow.
 */
import './style.css';
import { parseBankFile, parseFusionFile, normalizeName } from './parser.js';
import {
  loadDictionary, loadAllRecords, setMapping, setMappingBatch,
  deleteMapping, updateMapping, analyzeMapping, applyDictionaryToData,
  loadIgnoredSuggestions, ignoreSuggestion, unignoreSuggestion
} from './depara.js';
import { reconcile, computeStats } from './reconciliation.js';
import * as XLSX from 'xlsx';
import XLSXStyle from 'xlsx-js-style';

// ===== Application State =====
const state = {
  bankFiles: [],
  fusionFiles: [],
  bankData: [],
  fusionData: [],
  mappingAnalysis: null,
  reconciliationResults: [],
  valueSuggestions: [],
  currentStep: 1,
  sidebarOpen: false
};

// ===== DOM References =====
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

// Upload
const dropzoneBank = $('#dropzone-bank');
const dropzoneFusion = $('#dropzone-fusion');
const inputBank = $('#input-bank');
const inputFusion = $('#input-fusion');
const btnSelectBank = $('#btn-select-bank');
const btnSelectFusion = $('#btn-select-fusion');
const fileInfoBank = $('#file-info-bank');
const fileInfoFusion = $('#file-info-fusion');
const btnNextDepara = $('#btn-next-depara');

// De-Para
const deparaBody = $('#depara-body');
const deparaSearch = $('#depara-search');
const deparaEmpty = $('#depara-empty');
const btnRunConciliation = $('#btn-run-conciliation');
const btnBackUpload = $('#btn-back-upload');

// Results
const resultsBody = $('#results-body');
const suggestionsBody = $('#suggestions-body');
const valueSuggestionsSection = $('#value-suggestions-section');
const btnBackDepara = $('#btn-back-depara');
const btnNewAnalysis = $('#btn-new-analysis');
const btnExport = $('#btn-export');

// Sidebar
const sidebar = $('#sidebar-depara');
const sidebarOverlay = $('#sidebar-overlay');
const btnToggleSidebar = $('#btn-toggle-sidebar');
const btnCloseSidebar = $('#btn-close-sidebar');
const sidebarMappingsList = $('#sidebar-mappings-list');
const sidebarSearchInput = $('#sidebar-search-input');
const sidebarCount = $('#sidebar-count');
const sidebarNomeBanco = $('#sidebar-nome-banco');
const sidebarNomeFusion = $('#sidebar-nome-fusion');
const btnSidebarAdd = $('#btn-sidebar-add');
const sidebarIgnoredList = $('#sidebar-ignored-list');

// Loading
const loadingOverlay = $('#loading-overlay');
const loadingText = $('#loading-text');

// ===== Navigation =====
function goToStep(step) {
  state.currentStep = step;
  $$('.section').forEach(s => s.classList.remove('visible'));
  if (step === 1) $('#section-upload').classList.add('visible');
  if (step === 2) $('#section-depara').classList.add('visible');
  if (step === 3) $('#section-results').classList.add('visible');

  $$('.step').forEach(s => {
    const sn = parseInt(s.dataset.step);
    s.classList.remove('active', 'completed');
    if (sn === step) s.classList.add('active');
    if (sn < step) s.classList.add('completed');
  });
  $$('.step-line').forEach((line, i) => {
    line.classList.toggle('completed', i + 1 < step);
  });
}

// ===== Loading =====
function showLoading(text = 'Processando...') {
  loadingText.textContent = text;
  loadingOverlay.style.display = 'flex';
}
function hideLoading() {
  loadingOverlay.style.display = 'none';
}

// ===== Format Helpers =====
function formatBRL(value) {
  return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

// ===== Sidebar Logic =====
function openSidebar() {
  state.sidebarOpen = true;
  sidebar.classList.add('open');
  sidebarOverlay.classList.add('open');
  refreshSidebarMappings();
  refreshSidebarIgnored();
}

function closeSidebar() {
  state.sidebarOpen = false;
  sidebar.classList.remove('open');
  sidebarOverlay.classList.remove('open');
}

async function refreshSidebarMappings(searchTerm = '') {
  const records = await loadAllRecords();
  btnToggleSidebar.classList.toggle('has-mappings', records.length > 0);

  let filtered = records;
  if (searchTerm) {
    const term = searchTerm.toLowerCase();
    filtered = records.filter(r =>
      r.nome_banco.toLowerCase().includes(term) ||
      r.nome_fusion.toLowerCase().includes(term)
    );
  }

  sidebarCount.textContent = `${filtered.length} mapeamento${filtered.length !== 1 ? 's' : ''}`;
  sidebarMappingsList.innerHTML = '';

  if (filtered.length === 0) {
    sidebarMappingsList.innerHTML = `
      <div class="sidebar-empty">
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/>
        </svg>
        <p>Nenhum mapeamento ${searchTerm ? 'encontrado' : 'cadastrado'}</p>
        <span>${searchTerm ? 'Tente outro termo de busca' : 'Adicione mapeamentos usando o formulário acima'}</span>
      </div>`;
    return;
  }

  filtered.forEach(record => {
    const card = document.createElement('div');
    card.className = 'sidebar-mapping-card';
    card.dataset.id = record.id;
    card.innerHTML = `
      <div class="sidebar-mapping-info">
        <div class="sidebar-mapping-bank" title="${record.nome_banco}">${record.nome_banco}</div>
        <div class="sidebar-mapping-fusion" title="${record.nome_fusion}">${record.nome_fusion}</div>
        <div class="sidebar-mapping-date">${formatDate(record.criado_em)}</div>
      </div>
      <div class="sidebar-mapping-arrow">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>
      </div>
      <div class="sidebar-mapping-actions">
        <button class="sidebar-action-btn edit" title="Editar" data-action="edit" data-id="${record.id}">✎</button>
        <button class="sidebar-action-btn delete" title="Excluir" data-action="delete" data-id="${record.id}">✕</button>
      </div>`;
    sidebarMappingsList.appendChild(card);
  });
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  try {
    const d = new Date(dateStr);
    return d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' });
  } catch { return dateStr; }
}

async function refreshSidebarIgnored() {
  const records = await loadIgnoredSuggestions();
  sidebarIgnoredList.innerHTML = '';

  if (records.length === 0) {
    sidebarIgnoredList.innerHTML = `
      <div class="sidebar-empty">
        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M12 2v4"/><path d="m15.2 7.6 2.4-2.4"/><path d="M18 12h4"/><path d="m15.2 16.4 2.4 2.4"/><path d="M12 18v4"/><path d="m4.4 19.8 2.4-2.4"/><path d="M2 12h4"/><path d="m4.4 4.2 2.4 2.4"/>
        </svg>
        <p>Nenhuma recusa registrada</p>
        <span>Você ainda não recusou nenhuma sugestão de valor.</span>
      </div>`;
    return;
  }

  records.forEach(record => {
    const card = document.createElement('div');
    card.className = 'sidebar-mapping-card';
    card.innerHTML = `
      <div class="sidebar-mapping-info">
        <div class="sidebar-mapping-bank" title="${record.nome_banco_normalizado}">${record.nome_banco_normalizado}</div>
        <div class="sidebar-mapping-fusion" style="color: var(--accent-red)" title="${record.nome_fusion_normalizado}">${record.nome_fusion_normalizado}</div>
        <div class="sidebar-mapping-date">Recusado em: ${formatDate(record.criado_em)}</div>
      </div>
      <div class="sidebar-mapping-arrow">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
      </div>
      <div class="sidebar-mapping-actions">
        <button class="sidebar-action-btn delete unignore-btn" title="Desfazer Recusa" data-id="${record.id}">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
        </button>
      </div>`;
    
    card.querySelector('.unignore-btn').addEventListener('click', async () => {
      const success = await unignoreSuggestion(record.id);
      if (success) {
        refreshSidebarIgnored();
      } else {
        alert('Erro ao desfazer a recusa.');
      }
    });

    sidebarIgnoredList.appendChild(card);
  });
}

async function handleSidebarAdd() {
  const nomeBanco = sidebarNomeBanco.value.trim();
  const nomeFusion = sidebarNomeFusion.value.trim();
  if (!nomeBanco || !nomeFusion) { alert('Preencha ambos os campos.'); return; }

  const normalized = normalizeName(nomeBanco);
  const success = await setMapping(nomeBanco, normalized, nomeFusion);
  if (success) {
    sidebarNomeBanco.value = '';
    sidebarNomeFusion.value = '';
    await refreshSidebarMappings(sidebarSearchInput.value);
  } else {
    alert('Erro ao salvar o mapeamento.');
  }
}

async function handleSidebarAction(e) {
  const btn = e.target.closest('[data-action]');
  if (!btn) return;
  const action = btn.dataset.action;
  const id = btn.dataset.id;

  if (action === 'delete') {
    const success = await deleteMapping(id);
    if (success) {
      await refreshSidebarMappings(sidebarSearchInput.value);
      if (state.bankData.length > 0 && state.fusionData.length > 0) {
        state.mappingAnalysis = await analyzeMapping(state.bankData, state.fusionData);
        renderDePara(document.querySelector('.filter-btn.active')?.dataset.filter || 'all', deparaSearch.value);
      }
    } else {
      alert('Erro ao excluir: verifique a conexão com o servidor.');
    }
  }

  if (action === 'edit') {
    const card = btn.closest('.sidebar-mapping-card');
    const fusionEl = card.querySelector('.sidebar-mapping-fusion');
    const currentValue = fusionEl.textContent;
    const input = document.createElement('input');
    input.className = 'sidebar-edit-input';
    input.value = currentValue;
    fusionEl.replaceWith(input);
    input.focus();
    input.select();

    const save = async () => {
      const newValue = input.value.trim();
      if (newValue && newValue !== currentValue) {
        await updateMapping(id, newValue);
        if (state.bankData.length > 0 && state.fusionData.length > 0) {
          state.mappingAnalysis = await analyzeMapping(state.bankData, state.fusionData);
          renderDePara(document.querySelector('.filter-btn.active')?.dataset.filter || 'all', deparaSearch.value);
        }
      }
      await refreshSidebarMappings(sidebarSearchInput.value);
    };
    input.addEventListener('blur', save);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
      if (e.key === 'Escape') refreshSidebarMappings(sidebarSearchInput.value);
    });
  }
}

// ===== File Upload Logic =====
function setupDropzone(dropzone, input, type) {
  ['dragenter', 'dragover'].forEach(evt => {
    dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add('drag-over'); });
  });
  ['dragleave', 'drop'].forEach(evt => {
    dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove('drag-over'); });
  });
  dropzone.addEventListener('drop', (e) => {
    if (e.dataTransfer.files.length > 0) handleFileSelect(e.dataTransfer.files, type);
  });
  dropzone.addEventListener('click', (e) => {
    if (e.target.tagName !== 'BUTTON' && !e.target.closest('.btn-remove')) input.click();
  });
  input.addEventListener('change', () => {
    if (input.files.length > 0) handleFileSelect(input.files, type);
    input.value = ''; // Reset to allow re-selecting the same file
  });
}

function handleFileSelect(fileList, type) {
  let additions = 0;
  for (let i = 0; i < fileList.length; i++) {
    const file = fileList[i];
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      alert(`Arquivo ignorado: ${file.name} (Apenas .xlsx ou .xls são permitidos)`);
      continue;
    }
    if (type === 'bank') state.bankFiles.push(file);
    else state.fusionFiles.push(file);
    additions++;
  }
  
  if (additions > 0) {
    if (type === 'bank') dropzoneBank.classList.add('has-file');
    else dropzoneFusion.classList.add('has-file');
    renderFileList(type);
  }
}

function renderFileList(type) {
  const container = type === 'bank' ? fileInfoBank : fileInfoFusion;
  const files = type === 'bank' ? state.bankFiles : state.fusionFiles;
  
  container.innerHTML = '';
  if (files.length === 0) {
    container.style.display = 'none';
    if (type === 'bank') dropzoneBank.classList.remove('has-file');
    else dropzoneFusion.classList.remove('has-file');
  } else {
    container.style.display = 'flex';
    files.forEach((file, index) => {
      const badge = document.createElement('div');
      badge.className = `file-badge ${type}-badge`;
      badge.innerHTML = `
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        <span>${file.name}</span>
        <button class="btn-remove" data-index="${index}" title="Remover arquivo">✕</button>
      `;
      badge.querySelector('.btn-remove').addEventListener('click', (e) => {
        e.stopPropagation();
        removeFile(type, index);
      });
      container.appendChild(badge);
    });
  }
  updateUploadButton();
}

function removeFile(type, index) {
  if (type === 'bank') {
    state.bankFiles.splice(index, 1);
  } else {
    state.fusionFiles.splice(index, 1);
  }
  renderFileList(type);
}

function removeAllFiles(type) {
  if (type === 'bank') state.bankFiles = [];
  else state.fusionFiles = [];
  renderFileList(type);
}

function updateUploadButton() {
  btnNextDepara.disabled = !(state.bankFiles.length > 0 && state.fusionFiles.length > 0);
}

// ===== Parse Files & Go to De-Para =====
// ===== Parse Files & Go to De-Para =====
async function processFilesAndGoToDepara() {
  showLoading('Lendo e consolidando arquivos Excel...');
  try {
    // Banco
    let allBankData = [];
    for (const file of state.bankFiles) {
      const buffer = await file.arrayBuffer();
      const parsed = parseBankFile(buffer);
      allBankData = allBankData.concat(parsed);
    }
    state.bankData = allBankData;

    // Fusion
    let allFusionData = [];
    for (const file of state.fusionFiles) {
      const buffer = await file.arrayBuffer();
      const parsed = parseFusionFile(buffer);
      allFusionData = allFusionData.concat(parsed);
    }
    state.fusionData = allFusionData;

    if (state.bankData.length === 0) throw new Error('Nenhum dado encontrado nos arquivos do Banco.');
    if (state.fusionData.length === 0) throw new Error('Nenhum dado encontrado nos arquivos do Fusion.');

    state.mappingAnalysis = await analyzeMapping(state.bankData, state.fusionData);
    renderDePara();
    hideLoading();
    goToStep(2);
  } catch (error) {
    hideLoading();
    alert('Erro ao processar as planilhas:\n\n' + error.message);
    console.error(error);
  }
}

// ===== De-Para Rendering (Enhanced with entries & values) =====
function renderDePara(filter = 'all', searchTerm = '') {
  const { unmapped, mapped, autoMatched } = state.mappingAnalysis;
  const fusionNames = [...new Set(state.fusionData.map(d => d.supplier))].sort();

  // Build value summaries for each supplier
  const bankSummary = {};
  for (const e of state.bankData) {
    const key = e.supplierNormalized;
    if (!bankSummary[key]) bankSummary[key] = { count: 0, total: 0 };
    bankSummary[key].count++;
    bankSummary[key].total += e.value;
  }
  const fusionSummary = {};
  for (const e of state.fusionData) {
    const key = e.supplierNormalized;
    if (!fusionSummary[key]) fusionSummary[key] = { count: 0, total: 0 };
    fusionSummary[key].count++;
    fusionSummary[key].total += e.value;
  }

  const totalMapped = mapped.length + autoMatched.length;
  $('#stat-mapped').textContent = totalMapped;
  $('#stat-unmapped').textContent = unmapped.length;
  $('#stat-total').textContent = totalMapped + unmapped.length;

  deparaBody.innerHTML = '';

  let allEntries = [];
  unmapped.forEach(u => allEntries.push({ ...u, type: 'unmapped' }));
  autoMatched.forEach(a => allEntries.push({ ...a, type: 'auto' }));
  mapped.forEach(m => allEntries.push({ ...m, type: 'mapped' }));

  if (filter === 'unmapped') allEntries = allEntries.filter(e => e.type === 'unmapped');
  else if (filter === 'mapped') allEntries = allEntries.filter(e => e.type !== 'unmapped');

  if (searchTerm) {
    const term = searchTerm.toLowerCase();
    allEntries = allEntries.filter(e =>
      e.bankName.toLowerCase().includes(term) ||
      (e.fusionName && e.fusionName.toLowerCase().includes(term))
    );
  }

  if (allEntries.length === 0) {
    deparaEmpty.style.display = 'block';
    return;
  }
  deparaEmpty.style.display = 'none';

  allEntries.forEach(entry => {
    const tr = document.createElement('tr');

    // 1. Bank name (the abbreviated/different name)
    const tdBank = document.createElement('td');
    tdBank.textContent = entry.bankName;
    tdBank.style.fontWeight = '500';

    // 2. Bank entries/value
    const tdBankVal = document.createElement('td');
    tdBankVal.className = 'col-num';
    const bSummary = bankSummary[entry.bankNormalized];
    if (bSummary) {
      tdBankVal.innerHTML = `<span class="depara-value-info"><strong>${bSummary.count}</strong> lanç. · ${formatBRL(bSummary.total)}</span>`;
    } else {
      tdBankVal.innerHTML = '<span class="depara-value-info">—</span>';
    }

    // 3. Fusion name / select (the canonical name from Oracle)
    const tdFusion = document.createElement('td');
    // 4. Fusion entries/value
    const tdFusionVal = document.createElement('td');
    tdFusionVal.className = 'col-num';

    // Helper to update fusion value column
    function updateFusionValueDisplay(fusionName) {
      if (!fusionName) {
        tdFusionVal.innerHTML = '<span class="depara-value-info">—</span>';
        return;
      }
      const fKey = normalizeName(fusionName);
      const fSummary = fusionSummary[fKey];
      if (fSummary) {
        tdFusionVal.innerHTML = `<span class="depara-value-info"><strong>${fSummary.count}</strong> lanç. · ${formatBRL(fSummary.total)}</span>`;
      } else {
        tdFusionVal.innerHTML = '<span class="depara-value-info">—</span>';
      }
    }

    if (entry.type === 'unmapped') {
      const select = document.createElement('select');
      select.className = 'depara-select';
      const optEmpty = document.createElement('option');
      optEmpty.value = '';
      optEmpty.textContent = '— Selecione o fornecedor cadastrado —';
      select.appendChild(optEmpty);

      if (entry.suggestion) {
        const optSugg = document.createElement('option');
        optSugg.value = entry.suggestion;
        optSugg.textContent = `⭐ ${entry.suggestion} (sugestão)`;
        optSugg.selected = true;
        select.appendChild(optSugg);
      }

      fusionNames.forEach(fn => {
        if (fn !== entry.suggestion) {
          const opt = document.createElement('option');
          opt.value = fn;
          opt.textContent = fn;
          select.appendChild(opt);
        }
      });

      // Show Fusion values for the current selection (suggestion or empty)
      updateFusionValueDisplay(select.value || null);

      // Update values on dropdown change
      select.addEventListener('change', async () => {
        updateFusionValueDisplay(select.value || null);
        if (select.value) {
          await setMapping(entry.bankName, entry.bankNormalized, select.value);
          state.mappingAnalysis = await analyzeMapping(state.bankData, state.fusionData);
          renderDePara(
            document.querySelector('.filter-btn.active')?.dataset.filter || 'all',
            deparaSearch.value
          );
        }
      });
      tdFusion.appendChild(select);
    } else {
      tdFusion.textContent = entry.fusionName;
      tdFusion.style.color = 'var(--accent-emerald)';
      updateFusionValueDisplay(entry.fusionName);
    }

    // 5. Status
    const tdStatus = document.createElement('td');
    tdStatus.style.textAlign = 'center';
    const badge = document.createElement('span');
    badge.className = 'status-badge';
    if (entry.type === 'unmapped') { badge.classList.add('unmapped'); badge.textContent = 'Pendente'; }
    else if (entry.type === 'auto') { badge.classList.add('auto'); badge.textContent = 'Auto'; }
    else { badge.classList.add('mapped'); badge.textContent = 'Mapeado'; }
    tdStatus.appendChild(badge);

    tr.appendChild(tdBank);
    tr.appendChild(tdBankVal);
    tr.appendChild(tdFusion);
    tr.appendChild(tdFusionVal);
    tr.appendChild(tdStatus);
    deparaBody.appendChild(tr);
  });
}

// ===== Run Reconciliation =====
async function runReconciliation() {
  showLoading('Executando conciliação...');
  try {
    const ignoredSuggestions = await loadIgnoredSuggestions();
    const mappedBankData = await applyDictionaryToData(state.bankData);
    const { results, valueSuggestions } = reconcile(mappedBankData, state.fusionData, ignoredSuggestions);
    state.reconciliationResults = results;
    state.valueSuggestions = valueSuggestions;

    renderResults();
    renderValueSuggestions();
    hideLoading();
    goToStep(3);
  } catch (error) {
    hideLoading();
    alert('Erro na conciliação:\n\n' + error.message);
    console.error(error);
  }
}

// ===== Value Suggestions Rendering =====
function renderValueSuggestions() {
  const suggestions = state.valueSuggestions;
  suggestionsBody.innerHTML = '';

  if (suggestions.length === 0) {
    valueSuggestionsSection.style.display = 'none';
    return;
  }
  valueSuggestionsSection.style.display = 'block';

  suggestions.forEach((sugg, idx) => {
    const tr = document.createElement('tr');
    tr.className = sugg.userAction === 'approved' ? 'row-approved' :
                   sugg.userAction === 'rejected' ? 'row-rejected' : 'row-sugerido';
    tr.id = `sugg-row-${idx}`;

    const tdBank = document.createElement('td');
    tdBank.innerHTML = `<strong>${sugg.bankSupplier}</strong>`;

    const tdFusion = document.createElement('td');
    tdFusion.innerHTML = sugg.isPartialFusionGroup
      ? `<strong>${sugg.fusionSupplier}</strong><br><small style="color:var(--text-muted)">⚠ Lançamentos sem par dentro deste grupo</small>`
      : `<strong>${sugg.fusionSupplier}</strong>`;

    const tdValue = document.createElement('td');
    tdValue.className = 'col-num';
    if ((sugg.matchType === 'value_total' || sugg.matchType === 'value_total_partial') && sugg.difference === 0) {
      tdValue.innerHTML = `<span style="color: var(--accent-emerald)">${formatBRL(sugg.bankTotal)}</span>`;
    } else {
      tdValue.innerHTML = `
        <span>${formatBRL(sugg.bankTotal)}</span> / <span>${formatBRL(sugg.fusionTotal)}</span>
        ${sugg.valueMatches ? `<br><small style="color:var(--text-muted)">${sugg.valueMatches.length} valor(es) coincidente(s)</small>` : ''}
      `;
    }

    const tdActions = document.createElement('td');
    tdActions.className = 'col-status';

    if (sugg.userAction === 'approved') {
      tdActions.innerHTML = '<span class="result-status aprovado">✓ Aprovado</span>';
    } else if (sugg.userAction === 'rejected') {
      tdActions.innerHTML = '<span class="result-status recusado">✕ Recusado</span>';
    } else {
      tdActions.innerHTML = `
        <div class="action-btns">
          <button class="btn-approve" data-sugg-idx="${idx}" data-action="approve">✓ Aprovar</button>
          <button class="btn-reject" data-sugg-idx="${idx}" data-action="reject">✕ Recusar</button>
        </div>`;
    }

    tr.appendChild(tdBank);
    tr.appendChild(tdFusion);
    tr.appendChild(tdValue);
    tr.appendChild(tdActions);
    suggestionsBody.appendChild(tr);
  });
}

async function handleSuggestionAction(e) {
  const btn = e.target.closest('[data-sugg-idx]');
  if (!btn) return;

  const idx = parseInt(btn.dataset.suggIdx);
  const action = btn.dataset.action;
  const suggestion = state.valueSuggestions[idx];

  if (action === 'approve') {
    suggestion.userAction = 'approved';
    // Also save the mapping to SQLite for future use
    await setMapping(suggestion.bankSupplier, suggestion.bankSupplierNormalized, suggestion.fusionSupplier);
    await refreshSidebarMappings(sidebarSearchInput.value);
  } else if (action === 'reject') {
    suggestion.userAction = 'rejected';
    await ignoreSuggestion(suggestion.bankSupplierNormalized, suggestion.fusionSupplierNormalized);
    if (state.sidebarOpen) {
      refreshSidebarIgnored();
    }
  }

  renderValueSuggestions();
  updateResultStats();
}

function updateResultStats() {
  const stats = computeStats(state.reconciliationResults, state.valueSuggestions);
  $('#stat-conciliado').textContent = stats.conciliado;
  $('#stat-pendente').textContent = stats.pendente;
  $('#stat-sugerido').textContent = stats.sugerido;
  $('#stat-diff-total').textContent = formatBRL(stats.totalDifference);
}

// ===== Results Rendering =====
function renderResults(filter = 'all') {
  const results = state.reconciliationResults;
  updateResultStats();

  resultsBody.innerHTML = '';

  let filtered = results;
  if (filter === 'conciliado') {
    filtered = results.filter(r => r.status === 'conciliado' || r.userAction === 'approved');
  } else if (filter === 'pendente') {
    filtered = results.filter(r => r.status === 'pendente' && r.userAction !== 'approved');
  } else if (filter === 'sugerido') {
    // Show only value suggestions — the table is separate, but we can scroll to it
    valueSuggestionsSection.scrollIntoView({ behavior: 'smooth' });
    return;
  }

  // Sorting logic based on priority:
  // 1. Pendente com divergência (ambos > 0)
  // 2. Pendente SÓ no Banco
  // 3. Pendente SÓ no Fusion
  // 4. Conciliado
  filtered.sort((a, b) => {
    const getPriority = (item) => {
      if (item.status === 'conciliado' || item.userAction === 'approved') return 4;
      if (item.bankTotal > 0 && item.fusionTotal > 0) return 1; // Divergência
      if (item.bankTotal > 0 && (item.fusionTotal === 0 || !item.fusionTotal)) return 2; // Só Banco
      if ((item.bankTotal === 0 || !item.bankTotal) && item.fusionTotal > 0) return 3; // Só Fusion
      return 5;
    };

    const prioA = getPriority(a);
    const prioB = getPriority(b);
    
    if (prioA !== prioB) return prioA - prioB;
    
    // If same priority, sort by descending highest value
    const maxA = Math.max(a.bankTotal || 0, a.fusionTotal || 0);
    const maxB = Math.max(b.bankTotal || 0, b.fusionTotal || 0);
    return maxB - maxA;
  });

  filtered.forEach((result, idx) => {
    const tr = document.createElement('tr');
    tr.className = result.status === 'conciliado' ? 'row-conciliado' : 'row-pendente';

    const hasDetails = (result.bankEntries.length > 0 || result.fusionEntries.length > 0);

    // Expand
    const tdExpand = document.createElement('td');
    tdExpand.className = 'col-expand';
    if (hasDetails) {
      const btn = document.createElement('button');
      btn.className = 'expand-btn';
      btn.innerHTML = '▶';
      btn.addEventListener('click', () => {
        const detailRow = document.getElementById(`detail-${idx}`);
        if (detailRow) {
          detailRow.classList.toggle('visible');
          btn.classList.toggle('expanded');
        }
      });
      tdExpand.appendChild(btn);
    }

    // Supplier
    const tdSupplier = document.createElement('td');
    tdSupplier.textContent = result.supplier;
    tdSupplier.style.fontWeight = '500';
    if (result.onlyInBank) tdSupplier.title = 'Apenas no Banco';
    if (result.onlyInFusion) tdSupplier.title = 'Apenas no Fusion';

    // Bank total
    const tdBank = document.createElement('td');
    tdBank.className = 'col-num';
    tdBank.textContent = result.bankTotal > 0 ? formatBRL(result.bankTotal) : '—';

    // Fusion total
    const tdFusion = document.createElement('td');
    tdFusion.className = 'col-num';
    tdFusion.textContent = result.fusionTotal > 0 ? formatBRL(result.fusionTotal) : '—';

    // Status
    const tdStatus = document.createElement('td');
    tdStatus.className = 'col-status';
    const statusBadge = document.createElement('span');
    statusBadge.className = `result-status ${result.status}`;
    if (result.status === 'conciliado') {
      statusBadge.textContent = '✓ Conciliado';
    } else {
      statusBadge.textContent = '⚠ Pendente';
    }
    tdStatus.appendChild(statusBadge);

    // Difference
    const tdDiff = document.createElement('td');
    tdDiff.className = 'col-num';
    if (result.difference === 0) {
      tdDiff.textContent = '—';
      tdDiff.classList.add('value-zero');
    } else {
      tdDiff.textContent = formatBRL(result.difference);
      tdDiff.classList.add(result.difference > 0 ? 'value-negative' : 'value-positive');
    }

    tr.appendChild(tdExpand);
    tr.appendChild(tdSupplier);
    tr.appendChild(tdBank);
    tr.appendChild(tdFusion);
    tr.appendChild(tdStatus);
    tr.appendChild(tdDiff);
    resultsBody.appendChild(tr);

    // Detail row — show individual entries
    if (hasDetails) {
      const detailTr = document.createElement('tr');
      detailTr.className = 'detail-row';
      detailTr.id = `detail-${idx}`;

      const detailTd = document.createElement('td');
      detailTd.colSpan = 6;

      const detailContent = document.createElement('div');
      detailContent.className = 'detail-content';

      const detailGrid = document.createElement('div');
      detailGrid.className = 'detail-grid';

      // Bank column
      const bankCol = document.createElement('div');
      bankCol.className = 'detail-column';
      bankCol.innerHTML = `<h4>📄 Lançamentos Banco (${result.bankEntries.length})</h4>`;
      if (result.bankEntries.length === 0) {
        bankCol.innerHTML += '<div class="detail-item unmatched"><span>Nenhum lançamento</span><span>—</span></div>';
      }
      result.bankEntries.forEach(entry => {
        const matched = isEntryMatched(entry, result.matchDetails, 'bank');
        const div = document.createElement('div');
        div.className = `detail-item ${matched ? 'matched' : 'unmatched'}`;
        div.innerHTML = `
          <span>${entry.document || entry.supplier}</span>
          <span>${formatBRL(entry.value)}</span>`;
        bankCol.appendChild(div);
      });

      // Fusion column
      const fusionCol = document.createElement('div');
      fusionCol.className = 'detail-column';
      fusionCol.innerHTML = `<h4>🔗 Lançamentos Fusion (${result.fusionEntries.length})</h4>`;
      if (result.fusionEntries.length === 0) {
        fusionCol.innerHTML += '<div class="detail-item unmatched"><span>Nenhum lançamento</span><span>—</span></div>';
      }
      result.fusionEntries.forEach(entry => {
        const matched = isEntryMatched(entry, result.matchDetails, 'fusion');
        const div = document.createElement('div');
        div.className = `detail-item ${matched ? 'matched' : 'unmatched'}`;
        div.innerHTML = `
          <span>${entry.number || entry.supplier}</span>
          <span>${formatBRL(entry.value)}</span>`;
        fusionCol.appendChild(div);
      });

      detailGrid.appendChild(bankCol);
      detailGrid.appendChild(fusionCol);
      detailContent.appendChild(detailGrid);
      detailTd.appendChild(detailContent);
      detailTr.appendChild(detailTd);
      resultsBody.appendChild(detailTr);
    }
  });
}

function isEntryMatched(entry, matchDetails, source) {
  if (!matchDetails) return source === 'bank' ? true : true; // If no details, all considered matched (Level 1)
  if (!matchDetails.matches) return false;
  for (const match of matchDetails.matches) {
    const entries = source === 'bank' ? match.bank : match.fusion;
    if (entries.some(e => e === entry)) return true;
  }
  return false;
}

// ===== Export to Excel =====

// --- Style palette ---
const STYLE = {
  CURRENCY_FMT: '"R$" #,##0.00;[Red]-"R$" #,##0.00',
  CLIENT_HEADER: {
    font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 12 },
    fill: { patternType: 'solid', fgColor: { rgb: '1F2937' } },
    alignment: { vertical: 'center', horizontal: 'left' },
    border: { top: { style: 'thin', color: { rgb: '000000' } }, bottom: { style: 'thin', color: { rgb: '000000' } } }
  },
  STATUS_CONCILIADO: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '059669' } },
    alignment: { vertical: 'center', horizontal: 'center' }
  },
  STATUS_PENDENTE: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: 'D97706' } },
    alignment: { vertical: 'center', horizontal: 'center' }
  },
  TOTALS_LABEL: {
    font: { bold: true, color: { rgb: '111827' } },
    fill: { patternType: 'solid', fgColor: { rgb: 'E5E7EB' } },
    alignment: { vertical: 'center' }
  },
  TOTALS_BANK: {
    font: { bold: true, color: { rgb: '0C4A6E' } },
    fill: { patternType: 'solid', fgColor: { rgb: 'DBEAFE' } },
    alignment: { horizontal: 'right' },
    numFmt: '"R$" #,##0.00'
  },
  TOTALS_FUSION: {
    font: { bold: true, color: { rgb: '064E3B' } },
    fill: { patternType: 'solid', fgColor: { rgb: 'D1FAE5' } },
    alignment: { horizontal: 'right' },
    numFmt: '"R$" #,##0.00'
  },
  DIFF_LABEL: {
    font: { bold: true, color: { rgb: '7F1D1D' } },
    fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
    alignment: { horizontal: 'right' },
    numFmt: '"R$" #,##0.00'
  },
  HEADER_BANK: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '2563EB' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: { top: { style: 'thin', color: { rgb: '1E3A8A' } }, bottom: { style: 'thin', color: { rgb: '1E3A8A' } } }
  },
  HEADER_FUSION: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '059669' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: { top: { style: 'thin', color: { rgb: '064E3B' } }, bottom: { style: 'thin', color: { rgb: '064E3B' } } }
  },
  CELL_BANK: {
    fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
    border: { top: { style: 'thin', color: { rgb: 'BFDBFE' } }, bottom: { style: 'thin', color: { rgb: 'BFDBFE' } }, left: { style: 'thin', color: { rgb: 'BFDBFE' } }, right: { style: 'thin', color: { rgb: 'BFDBFE' } } }
  },
  CELL_BANK_MONEY: {
    fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
    font: { color: { rgb: '1E3A8A' } },
    alignment: { horizontal: 'right' },
    numFmt: '"R$" #,##0.00',
    border: { top: { style: 'thin', color: { rgb: 'BFDBFE' } }, bottom: { style: 'thin', color: { rgb: 'BFDBFE' } }, left: { style: 'thin', color: { rgb: 'BFDBFE' } }, right: { style: 'thin', color: { rgb: 'BFDBFE' } } }
  },
  CELL_FUSION: {
    fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
    border: { top: { style: 'thin', color: { rgb: 'A7F3D0' } }, bottom: { style: 'thin', color: { rgb: 'A7F3D0' } }, left: { style: 'thin', color: { rgb: 'A7F3D0' } }, right: { style: 'thin', color: { rgb: 'A7F3D0' } } }
  },
  CELL_FUSION_MONEY: {
    fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
    font: { color: { rgb: '064E3B' } },
    alignment: { horizontal: 'right' },
    numFmt: '"R$" #,##0.00',
    border: { top: { style: 'thin', color: { rgb: 'A7F3D0' } }, bottom: { style: 'thin', color: { rgb: 'A7F3D0' } }, left: { style: 'thin', color: { rgb: 'A7F3D0' } }, right: { style: 'thin', color: { rgb: 'A7F3D0' } } }
  },
  CELL_EMPTY_BANK: { fill: { patternType: 'solid', fgColor: { rgb: 'F8FAFC' } } },
  CELL_EMPTY_FUSION: { fill: { patternType: 'solid', fgColor: { rgb: 'F8FAFC' } } }
};

/**
 * Split each reconciliation result into two buckets:
 *   - matched pairs (go to Conciliados sheet, even if the supplier is PENDENTE overall)
 *   - unmatched entries (go to Pendentes sheet)
 * Level 1 (full match) → fully conciliado.
 * Level 4 (only bank or only fusion) → fully pendente.
 */
function splitResultsByMatchState(results) {
  const conciliadosBlocks = [];
  const pendentesBlocks = [];

  const sum = (arr) => arr.reduce((s, e) => s + (e.value || 0), 0);

  for (const r of results) {
    const isFullyConciliado = r.status === 'conciliado' || r.userAction === 'approved';

    if (isFullyConciliado) {
      conciliadosBlocks.push({
        supplier: r.supplier,
        bankTotal: r.bankTotal || 0,
        fusionTotal: r.fusionTotal || 0,
        difference: r.difference || 0,
        statusLabel: 'CONCILIADO',
        statusKind: 'conciliado',
        bankEntries: r.bankEntries || [],
        fusionEntries: r.fusionEntries || []
      });
      continue;
    }

    // PENDENTE result — may have partial matches in matchDetails
    if (r.matchDetails && r.matchDetails.matches && r.matchDetails.matches.length > 0) {
      const matchedBank = [];
      const matchedFusion = [];
      for (const m of r.matchDetails.matches) {
        matchedBank.push(...m.bank);
        matchedFusion.push(...m.fusion);
      }

      if (matchedBank.length > 0 || matchedFusion.length > 0) {
        const mbTotal = sum(matchedBank);
        const mfTotal = sum(matchedFusion);
        conciliadosBlocks.push({
          supplier: r.supplier,
          bankTotal: mbTotal,
          fusionTotal: mfTotal,
          difference: mbTotal - mfTotal,
          statusLabel: 'CONCILIADO (parcial)',
          statusKind: 'conciliado',
          bankEntries: matchedBank,
          fusionEntries: matchedFusion
        });
      }

      const ub = r.matchDetails.unmatchedBank || [];
      const uf = r.matchDetails.unmatchedFusion || [];
      if (ub.length > 0 || uf.length > 0) {
        const ubTotal = sum(ub);
        const ufTotal = sum(uf);
        pendentesBlocks.push({
          supplier: r.supplier,
          bankTotal: ubTotal,
          fusionTotal: ufTotal,
          difference: ubTotal - ufTotal,
          statusLabel: 'PENDENTE',
          statusKind: 'pendente',
          bankEntries: ub,
          fusionEntries: uf
        });
      }
    } else {
      // No partial matches — everything pending as is
      pendentesBlocks.push({
        supplier: r.supplier,
        bankTotal: r.bankTotal || 0,
        fusionTotal: r.fusionTotal || 0,
        difference: r.difference || 0,
        statusLabel: 'PENDENTE',
        statusKind: 'pendente',
        bankEntries: r.bankEntries || [],
        fusionEntries: r.fusionEntries || []
      });
    }
  }

  conciliadosBlocks.sort((a, b) => a.supplier.localeCompare(b.supplier));
  pendentesBlocks.sort((a, b) => a.supplier.localeCompare(b.supplier));
  return { conciliadosBlocks, pendentesBlocks };
}

/**
 * Build a styled per-client sheet.
 * Columns: A=Documento Banco, B=Valor Banco, C=Documento Fusion, D=Valor Fusion, E=Status/Diferença
 */
function buildStyledClientSheet(blocks) {
  // Build the AOA first (strings & numbers), we'll then stamp styles onto specific cells.
  const aoa = [];
  const styleMap = {}; // cellRef -> style object
  const merges = [];

  const setStyle = (row, col, style) => {
    const ref = XLSX.utils.encode_cell({ r: row, c: col });
    styleMap[ref] = style;
  };

  blocks.forEach((b, idx) => {
    if (idx > 0) aoa.push(['', '', '', '', '']); // blank separator

    const clientRow = aoa.length;
    aoa.push([`CLIENTE: ${b.supplier}`, '', '', '', b.statusLabel]);
    merges.push({ s: { r: clientRow, c: 0 }, e: { r: clientRow, c: 3 } });
    setStyle(clientRow, 0, STYLE.CLIENT_HEADER);
    setStyle(clientRow, 1, STYLE.CLIENT_HEADER);
    setStyle(clientRow, 2, STYLE.CLIENT_HEADER);
    setStyle(clientRow, 3, STYLE.CLIENT_HEADER);
    setStyle(clientRow, 4, b.statusKind === 'conciliado' ? STYLE.STATUS_CONCILIADO : STYLE.STATUS_PENDENTE);

    const totalsRow = aoa.length;
    aoa.push(['Total Banco', b.bankTotal, 'Total Fusion', b.fusionTotal, b.difference]);
    setStyle(totalsRow, 0, STYLE.TOTALS_LABEL);
    setStyle(totalsRow, 1, STYLE.TOTALS_BANK);
    setStyle(totalsRow, 2, STYLE.TOTALS_LABEL);
    setStyle(totalsRow, 3, STYLE.TOTALS_FUSION);
    setStyle(totalsRow, 4, STYLE.DIFF_LABEL);

    const headerRow = aoa.length;
    aoa.push(['Documento Banco', 'Valor Banco', 'Documento Fusion', 'Valor Fusion', '']);
    setStyle(headerRow, 0, STYLE.HEADER_BANK);
    setStyle(headerRow, 1, STYLE.HEADER_BANK);
    setStyle(headerRow, 2, STYLE.HEADER_FUSION);
    setStyle(headerRow, 3, STYLE.HEADER_FUSION);

    const bank = b.bankEntries || [];
    const fusion = b.fusionEntries || [];
    const max = Math.max(bank.length, fusion.length, 1);
    for (let i = 0; i < max; i++) {
      const rowIdx = aoa.length;
      const be = bank[i];
      const fe = fusion[i];
      aoa.push([
        be ? (be.document || be.supplier || '') : '',
        be ? be.value : '',
        fe ? (fe.number || fe.document || fe.supplier || '') : '',
        fe ? fe.value : '',
        ''
      ]);
      setStyle(rowIdx, 0, be ? STYLE.CELL_BANK : STYLE.CELL_EMPTY_BANK);
      setStyle(rowIdx, 1, be ? STYLE.CELL_BANK_MONEY : STYLE.CELL_EMPTY_BANK);
      setStyle(rowIdx, 2, fe ? STYLE.CELL_FUSION : STYLE.CELL_EMPTY_FUSION);
      setStyle(rowIdx, 3, fe ? STYLE.CELL_FUSION_MONEY : STYLE.CELL_EMPTY_FUSION);
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 40 }, { wch: 18 }, { wch: 40 }, { wch: 18 }, { wch: 22 }];
  if (merges.length) ws['!merges'] = merges;

  // Stamp styles onto cells (cell must exist in ws for styles to be written)
  Object.entries(styleMap).forEach(([ref, style]) => {
    const cell = ws[ref];
    if (!cell) {
      // Empty cell — create a blank one so xlsx-js-style writes the style
      ws[ref] = { t: 's', v: '', s: style };
    } else {
      cell.s = style;
      // Apply currency format explicitly when the style carries numFmt
      if (style.numFmt && typeof cell.v === 'number') {
        cell.z = style.numFmt;
        cell.t = 'n';
      }
    }
  });

  return ws;
}

function exportResults() {
  const results = state.reconciliationResults;
  const suggestions = state.valueSuggestions || [];

  const { conciliadosBlocks, pendentesBlocks } = splitResultsByMatchState(results);

  const wb = XLSXStyle.utils.book_new();

  if (pendentesBlocks.length > 0) {
    const ws = buildStyledClientSheet(pendentesBlocks);
    XLSXStyle.utils.book_append_sheet(wb, ws, 'Pendentes');
  }

  if (conciliadosBlocks.length > 0) {
    const ws = buildStyledClientSheet(conciliadosBlocks);
    XLSXStyle.utils.book_append_sheet(wb, ws, 'Conciliados');
  }

  if (suggestions.length > 0) {
    const suggBlocks = suggestions.map(s => {
      const bTotal = (s.bankEntries || []).reduce((t, e) => t + (e.value || 0), 0);
      const fTotal = (s.fusionEntries || []).reduce((t, e) => t + (e.value || 0), 0);
      const isApproved = s.userAction === 'approved';
      return {
        supplier: `${s.bankSupplier}  →  ${s.fusionSupplier}`,
        bankTotal: bTotal,
        fusionTotal: fTotal,
        difference: bTotal - fTotal,
        statusLabel: isApproved ? 'APROVADO' : (s.userAction === 'rejected' ? 'RECUSADO' : 'SUGERIDO'),
        statusKind: isApproved ? 'conciliado' : 'pendente',
        bankEntries: s.bankEntries || [],
        fusionEntries: s.fusionEntries || []
      };
    });
    const ws = buildStyledClientSheet(suggBlocks);
    XLSXStyle.utils.book_append_sheet(wb, ws, 'Sugestões');
  }

  if (pendentesBlocks.length === 0 && conciliadosBlocks.length === 0 && suggestions.length === 0) {
    const wsEmpty = XLSXStyle.utils.aoa_to_sheet([['Nenhum dado para exportar']]);
    XLSXStyle.utils.book_append_sheet(wb, wsEmpty, 'Vazio');
  }

  XLSXStyle.writeFile(wb, `conciliacao_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// ===== Event Listeners =====
async function init() {
  // File upload
  setupDropzone(dropzoneBank, inputBank, 'bank');
  setupDropzone(dropzoneFusion, inputFusion, 'fusion');
  btnSelectBank.addEventListener('click', (e) => { e.stopPropagation(); inputBank.click(); });
  btnSelectFusion.addEventListener('click', (e) => { e.stopPropagation(); inputFusion.click(); });
  // Navigation
  btnNextDepara.addEventListener('click', processFilesAndGoToDepara);
  btnBackUpload.addEventListener('click', () => goToStep(1));
  btnBackDepara.addEventListener('click', () => goToStep(2));
  btnNewAnalysis.addEventListener('click', () => {
    state.bankData = [];
    state.fusionData = [];
    state.reconciliationResults = [];
    state.valueSuggestions = [];
    removeAllFiles('bank');
    removeAllFiles('fusion');
    goToStep(1);
  });

  // Run
  btnRunConciliation.addEventListener('click', runReconciliation);

  // De-Para filters
  $$('.filter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      $$('.filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderDePara(btn.dataset.filter, deparaSearch.value);
    });
  });
  deparaSearch.addEventListener('input', () => {
    const activeFilter = document.querySelector('.filter-btn.active')?.dataset.filter || 'all';
    renderDePara(activeFilter, deparaSearch.value);
  });

  // Results filters
  $$('.rfilter-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      $$('.rfilter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderResults(btn.dataset.rfilter);
    });
  });

  // Export
  btnExport.addEventListener('click', exportResults);

  // Sidebar
  btnToggleSidebar.addEventListener('click', () => state.sidebarOpen ? closeSidebar() : openSidebar());
  btnCloseSidebar.addEventListener('click', closeSidebar);
  sidebarOverlay.addEventListener('click', closeSidebar);
  btnSidebarAdd.addEventListener('click', handleSidebarAdd);
  sidebarNomeFusion.addEventListener('keydown', (e) => { if (e.key === 'Enter') handleSidebarAdd(); });
  let sidebarSearchTimeout;
  sidebarSearchInput.addEventListener('input', () => {
    clearTimeout(sidebarSearchTimeout);
    sidebarSearchTimeout = setTimeout(() => refreshSidebarMappings(sidebarSearchInput.value), 200);
  });
  sidebarMappingsList.addEventListener('click', handleSidebarAction);

  // Sidebar Tabs
  $$('.sidebar-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      $$('.sidebar-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      $$('.sidebar-tab-pane').forEach(p => p.classList.remove('active'));
      $(`#sidebar-content-${tab.dataset.sidebarTab}`).classList.add('active');
    });
  });
  // Value suggestions approve/reject
  suggestionsBody.addEventListener('click', handleSuggestionAction);

  // Keyboard
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && state.sidebarOpen) closeSidebar();
  });

  // Check existing mappings
  const existingRecords = await loadAllRecords();
  btnToggleSidebar.classList.toggle('has-mappings', existingRecords.length > 0);
}

init();
