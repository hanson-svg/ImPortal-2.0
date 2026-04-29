// ── File input display ────────────────────────────────────────────────────────
const fileInput = document.getElementById('fileInput');
const fileDisplay = document.getElementById('fileDisplay');

fileInput.addEventListener('change', () => {
  const file = fileInput.files[0];
  if (file) {
    fileDisplay.innerHTML = `
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
      </svg>
      ${file.name}
    `;
    fileDisplay.classList.add('has-file');
  }
});

// ── Utilities ─────────────────────────────────────────────────────────────────
function setStatus(type, message, showSpinner = false) {
  const bar     = document.getElementById('statusBar');
  const text    = document.getElementById('statusText');
  const spinner = document.getElementById('statusSpinner');
  bar.className = `status-bar visible ${type}`;
  text.textContent = message;
  spinner.classList.toggle('hidden', !showSpinner);
}

function renderPreviewTable(rows, containerId) {
  const container = document.getElementById(containerId);
  if (!rows || rows.length === 0) {
    container.innerHTML = '<div class="empty-state">No data to preview</div>';
    return;
  }
  const keys = Object.keys(rows[0]);
  container.innerHTML = `
    <table>
      <thead><tr>${keys.map(k => `<th title="${k}">${k}</th>`).join('')}</tr></thead>
      <tbody>
        ${rows.map(row =>
          `<tr>${keys.map(k => `<td title="${String(row[k] ?? '')}">${row[k] ?? ''}</td>`).join('')}</tr>`
        ).join('')}
      </tbody>
    </table>
  `;
}

const GIS_COLS = new Set(['PhaseID', 'LotNum', 'BlockNum', 'ST_NUM', 'ST_NAME']);

// ── Tableau connection ────────────────────────────────────────────────────────
const TABLEAU_SERVER        = 'https://10ay.online.tableau.com';
const TABLEAU_SITE          = 'betenboughhomes';
const TABLEAU_API           = `${TABLEAU_SERVER}/api/3.22`;
const PHASE_NAME_COL        = 'Phase';
const PHASE_ID_COL          = 'Phase ID';
const TABLEAU_VIEW_URL_NAME = 'PhaseList';

function rowKey(attrs) {
  const lot   = String(attrs.LotNum   || '').trim();
  const block = String(attrs.BlockNum || '').trim();
  return block ? `${lot}|${block}` : lot;
}

// ── SVG map state ─────────────────────────────────────────────────────────────
let svgEl           = null;
let svgActiveGroup  = null;
let svgMatchedKeys  = new Set();
let svgContentGroup = null;
let svgTransform    = { x: 0, y: 0, s: 1 };

// Handing assignment state
let handingMap    = new Map();   // key → 'Left' | 'Right' | ''
let activeHanding = '';          // currently active paint mode

function applyMapTransform() {
  if (svgContentGroup) {
    const { x, y, s } = svgTransform;
    svgContentGroup.setAttribute('transform', `translate(${x},${y}) scale(${s})`);
  }
}

function styleFeatureGroup(fg, state) {
  const key     = fg.dataset.key;
  const matched = svgMatchedKeys.has(key);
  const handing = handingMap.get(key) || '';

  fg.querySelectorAll('path').forEach(p => {
    if (!matched) {
      p.setAttribute('fill', '#e4e2bd'); p.setAttribute('fill-opacity', '0.35');
      p.setAttribute('stroke', '#a0a87a'); p.setAttribute('stroke-width', '1');
      return;
    }
    // Resolve fill/stroke by handing
    const fill   = handing === 'Left'  ? '#ede9fe'
                 : handing === 'Right' ? '#fee2e2'
                 : '#73b4ae';
    const stroke = handing === 'Left'  ? '#7c3aed'
                 : handing === 'Right' ? '#dc2626'
                 : '#375957';

    if (state === 'selected') {
      p.setAttribute('fill', handing === 'Left'  ? '#ddd6fe'
                           : handing === 'Right' ? '#fecaca'
                           : '#db5d00');
      p.setAttribute('fill-opacity', '0.82');
      p.setAttribute('stroke', handing === 'Left'  ? '#5b21b6'
                              : handing === 'Right' ? '#b91c1c'
                              : '#a34500');
      p.setAttribute('stroke-width', '2.5');
    } else if (state === 'hover') {
      p.setAttribute('fill', fill); p.setAttribute('fill-opacity', '0.78');
      p.setAttribute('stroke', stroke); p.setAttribute('stroke-width', '1.8');
    } else {
      p.setAttribute('fill', fill);
      p.setAttribute('fill-opacity', handing ? '0.68' : '0.5');
      p.setAttribute('stroke', stroke); p.setAttribute('stroke-width', '1.2');
    }
  });

  // Lot label color
  fg.querySelectorAll('text').forEach(t =>
    t.setAttribute('fill',
      handing === 'Left'  ? '#5b21b6' :
      handing === 'Right' ? '#991b1b' : '#1a2725'
    )
  );
}

function updateMapFeatureStyle(key) {
  if (!svgEl) return;
  const fg = svgEl.querySelector(`g[data-key="${key}"]`);
  if (fg) styleFeatureGroup(fg, fg === svgActiveGroup ? 'selected' : 'default');
}

function syncTableHanding(key, handing) {
  const wrap = document.getElementById('editTableWrap');
  if (!wrap) return;
  const tr = wrap.querySelector(`tr[data-key="${key}"]`);
  if (!tr) return;
  const sel = tr.querySelector('select[data-col="Handing"]');
  if (sel) {
    sel.value = handing;
    sel.closest('td').classList.toggle('cell-edited', handing !== '');
  }
}

function selectMapByKey(key) {
  if (!svgEl) return;
  const fg = svgEl.querySelector(`g[data-key="${key}"]`);
  if (!fg) return;
  if (svgActiveGroup && svgActiveGroup !== fg) styleFeatureGroup(svgActiveGroup, 'default');
  svgActiveGroup = fg;
  styleFeatureGroup(fg, 'selected');
}

function renderSVGMap(features, blended) {
  const container = document.getElementById('map');

  // Clean up previous window listeners
  if (svgEl?._cleanup) svgEl._cleanup();
  container.innerHTML = '';
  svgActiveGroup = null;
  svgTransform   = { x: 0, y: 0, s: 1 };
  svgMatchedKeys = new Set(blended.map(rowKey));

  // Collect all coordinate points to derive the bounding box
  const allPts = features.flatMap(f => (f.geometry?.rings || []).flat());
  if (!allPts.length) {
    container.innerHTML = '<div class="empty-state" style="padding-top:60px">No geometry available for this Phase ID</div>';
    return null;
  }

  const minLon = Math.min(...allPts.map(p => p[0]));
  const maxLon = Math.max(...allPts.map(p => p[0]));
  const minLat = Math.min(...allPts.map(p => p[1]));
  const maxLat = Math.max(...allPts.map(p => p[1]));

  // Use CSS-declared dimensions; fall back to sensible defaults if not yet laid out
  const W   = container.clientWidth  || 400;
  const H   = container.clientHeight || 450;
  const PAD = 28;

  const lonSpan = maxLon - minLon || 0.001;
  const latSpan = maxLat - minLat || 0.001;
  const scale   = Math.min((W - PAD * 2) / lonSpan, (H - PAD * 2) / latSpan);
  const offX    = (W - lonSpan * scale) / 2;
  const offY    = (H - latSpan * scale) / 2;

  const project = ([lon, lat]) => [
    offX + (lon - minLon) * scale,
    H - (offY + (lat - minLat) * scale)   // flip Y axis
  ];

  const NS  = 'http://www.w3.org/2000/svg';
  const svg = document.createElementNS(NS, 'svg');
  svg.setAttribute('width',  W);
  svg.setAttribute('height', H);
  svg.style.cssText = 'display:block;cursor:grab;user-select:none;';

  // Background
  const bg = document.createElementNS(NS, 'rect');
  bg.setAttribute('width', W); bg.setAttribute('height', H);
  bg.setAttribute('fill', '#f0f5f4');
  svg.appendChild(bg);

  // Pan/zoom wrapper
  const contentG = document.createElementNS(NS, 'g');
  svg.appendChild(contentG);
  svgContentGroup = contentG;

  // ── Draw features ──────────────────────────────────────────────────────────
  features.forEach(f => {
    if (!f.geometry?.rings?.length) return;
    const key     = rowKey(f.attributes);
    const matched = svgMatchedKeys.has(key);

    const fg = document.createElementNS(NS, 'g');
    fg.dataset.key  = key;
    fg.style.cursor = matched ? 'pointer' : 'default';

    // Polygon paths
    f.geometry.rings.forEach(ring => {
      const d = ring.map(project)
        .map(([x, y], i) => `${i ? 'L' : 'M'}${x.toFixed(1)},${y.toFixed(1)}`)
        .join('') + 'Z';
      const path = document.createElementNS(NS, 'path');
      path.setAttribute('d', d);
      fg.appendChild(path);
    });
    styleFeatureGroup(fg, 'default');

    // Lot label at ring centroid
    const ring0 = f.geometry.rings[0];
    if (ring0?.length) {
      const cx = ring0.reduce((s, p) => s + p[0], 0) / ring0.length;
      const cy = ring0.reduce((s, p) => s + p[1], 0) / ring0.length;
      const [sx, sy] = project([cx, cy]);
      const txt = document.createElementNS(NS, 'text');
      txt.setAttribute('x', sx); txt.setAttribute('y', sy);
      txt.setAttribute('text-anchor', 'middle');
      txt.setAttribute('dominant-baseline', 'middle');
      txt.setAttribute('font-size', '8');
      txt.setAttribute('font-family', '-apple-system, BlinkMacSystemFont, sans-serif');
      txt.setAttribute('fill', '#1a2725');
      txt.setAttribute('pointer-events', 'none');
      txt.textContent = String(f.attributes.LotNum || '');
      fg.appendChild(txt);
    }

    // Hover
    fg.addEventListener('mouseenter', () => {
      if (fg !== svgActiveGroup) styleFeatureGroup(fg, 'hover');
    });
    fg.addEventListener('mouseleave', () => {
      if (fg !== svgActiveGroup) styleFeatureGroup(fg, 'default');
    });

    // Click — assignment paint or navigate/select
    if (matched) {
      fg.addEventListener('click', e => {
        if (dragged) return;
        e.stopPropagation();
        if (activeHanding !== '') {
          // Paint mode: assign handing immediately, no selection change needed
          handingMap.set(key, activeHanding);
          styleFeatureGroup(fg, fg === svgActiveGroup ? 'selected' : 'default');
          syncTableHanding(key, activeHanding);
        } else {
          // Navigate mode: select + jump to table row
          if (svgActiveGroup && svgActiveGroup !== fg) styleFeatureGroup(svgActiveGroup, 'default');
          svgActiveGroup = fg;
          styleFeatureGroup(fg, 'selected');
          highlightTableRow(key);
        }
      });
    }

    contentG.appendChild(fg);
  });

  // ── Pan / zoom ─────────────────────────────────────────────────────────────
  let drag    = null;
  let dragged = false;

  const onWheel = e => {
    e.preventDefault();
    const factor = e.deltaY < 0 ? 1.12 : 1 / 1.12;
    const rect   = svg.getBoundingClientRect();
    const mx     = e.clientX - rect.left;
    const my     = e.clientY - rect.top;
    svgTransform.x = mx - (mx - svgTransform.x) * factor;
    svgTransform.y = my - (my - svgTransform.y) * factor;
    svgTransform.s *= factor;
    applyMapTransform();
  };

  const onMouseDown = e => {
    drag    = { cx: e.clientX, cy: e.clientY, tx: svgTransform.x, ty: svgTransform.y };
    dragged = false;
    svg.style.cursor = 'grabbing';
  };

  const onMouseMove = e => {
    if (!drag) return;
    const dx = e.clientX - drag.cx;
    const dy = e.clientY - drag.cy;
    if (Math.abs(dx) > 4 || Math.abs(dy) > 4) dragged = true;
    svgTransform.x = drag.tx + dx;
    svgTransform.y = drag.ty + dy;
    applyMapTransform();
  };

  const onMouseUp = () => {
    drag = null;
    svg.style.cursor = 'grab';
  };

  svg.addEventListener('wheel', onWheel, { passive: false });
  svg.addEventListener('mousedown', onMouseDown);
  window.addEventListener('mousemove', onMouseMove);
  window.addEventListener('mouseup', onMouseUp);

  // Remove window listeners when the map is replaced
  svg._cleanup = () => {
    window.removeEventListener('mousemove', onMouseMove);
    window.removeEventListener('mouseup', onMouseUp);
  };

  svgEl = svg;
  container.appendChild(svg);
  return svg;
}

// ── Table highlight & editor ───────────────────────────────────────────────────
let blendedData = [];

function highlightTableRow(key) {
  const wrap = document.getElementById('editTableWrap');
  wrap.querySelectorAll('tr[data-key]').forEach(r => r.classList.remove('row-selected'));
  const target = wrap.querySelector(`tr[data-key="${key}"]`);
  if (target) {
    target.classList.add('row-selected');
    target.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

function renderEditTable(blended) {
  const wrap = document.getElementById('editTableWrap');
  if (!blended.length) {
    wrap.innerHTML = '<div class="empty-state">No blended rows to display</div>';
    return;
  }

  const keys = Object.keys(blended[0]);
  wrap.innerHTML = `
    <table class="edit-table">
      <thead>
        <tr>${keys.map(k =>
          `<th class="${GIS_COLS.has(k) ? 'col-gis' : ''}" title="${k}">${k}</th>`
        ).join('')}</tr>
      </thead>
      <tbody>
        ${blended.map((row, idx) => `
          <tr data-key="${rowKey(row)}" data-idx="${idx}">
            ${keys.map(k => {
              if (GIS_COLS.has(k)) {
                return `<td class="col-gis" title="${row[k] ?? ''}">${row[k] ?? ''}</td>`;
              } else if (k === 'Premium') {
                return `<td class="col-premium">
                  <div class="premium-wrapper">
                    <span class="dollar-sign">$</span>
                    <input type="number" class="premium-input" data-col="Premium"
                           min="0" step="1" value="${row[k] ?? ''}" placeholder="0">
                  </div>
                </td>`;
              } else if (k === 'Handing') {
                return `<td class="col-handing">
                  <select class="handing-select" data-col="Handing">
                    <option value="" ${!row[k] ? 'selected' : ''}>—</option>
                    <option value="Left"  ${row[k] === 'Left'  ? 'selected' : ''}>Left</option>
                    <option value="Right" ${row[k] === 'Right' ? 'selected' : ''}>Right</option>
                  </select>
                </td>`;
              } else {
                return `<td class="col-sheet" contenteditable="true"
                            data-col="${k}"
                            data-original="${String(row[k] ?? '').replace(/"/g, '&quot;')}"
                            title="${String(row[k] ?? '')}">${row[k] ?? ''}</td>`;
              }
            }).join('')}
          </tr>`
        ).join('')}
      </tbody>
    </table>
  `;

  // Highlight edited cells — contenteditable columns
  wrap.querySelectorAll('td.col-sheet').forEach(cell => {
    cell.addEventListener('input', () => {
      cell.classList.toggle('cell-edited', cell.textContent.trim() !== cell.dataset.original);
    });
  });

  // Premium — mark cell edited
  wrap.querySelectorAll('.premium-input').forEach(el => {
    el.addEventListener('change', () => el.closest('td').classList.add('cell-edited'));
    el.addEventListener('input',  () => el.closest('td').classList.add('cell-edited'));
  });

  // Handing — mark cell edited + sync map color
  wrap.querySelectorAll('.handing-select').forEach(sel => {
    sel.addEventListener('change', () => {
      const key = sel.closest('tr').dataset.key;
      handingMap.set(key, sel.value);
      updateMapFeatureStyle(key);
      sel.closest('td').classList.toggle('cell-edited', sel.value !== '');
    });
  });

  // Click a GIS cell / row → select on map
  wrap.querySelectorAll('tr[data-key]').forEach(tr => {
    tr.addEventListener('click', e => {
      if (e.target.classList.contains('col-sheet')) return;
      selectMapByKey(tr.dataset.key);
      tr.classList.add('row-selected');
    });
  });
}

function getEditedData() {
  const wrap = document.getElementById('editTableWrap');
  if (!wrap || !blendedData.length) return blendedData;
  return blendedData.map((row, idx) => {
    const tr = wrap.querySelector(`tr[data-idx="${idx}"]`);
    if (!tr) return row;
    const edited = { ...row };
    // Contenteditable cells
    tr.querySelectorAll('td[data-col]').forEach(cell => {
      edited[cell.dataset.col] = cell.textContent.trim();
    });
    // Premium — store as integer (empty string if blank)
    tr.querySelectorAll('input[data-col]').forEach(input => {
      edited[input.dataset.col] = input.value === '' ? '' : parseInt(input.value, 10);
    });
    // Handing — store as string
    tr.querySelectorAll('select[data-col]').forEach(sel => {
      edited[sel.dataset.col] = sel.value;
    });
    return edited;
  });
}

function initMapAndEditor(features, blended) {
  blendedData = blended;
  handingMap   = new Map();   // reset assignments for new dataset
  renderEditTable(blended);
  setTimeout(() => renderSVGMap(features, blended), 0);
}

// ── Tableau API ───────────────────────────────────────────────────────────────
async function tableauSignIn(patName, patSecret) {
  const res = await fetch(`${TABLEAU_API}/auth/signin`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
    body: JSON.stringify({
      credentials: {
        personalAccessTokenName: patName,
        personalAccessTokenSecret: patSecret,
        site: { contentUrl: TABLEAU_SITE }
      }
    })
  });
  if (!res.ok) throw new Error(`Tableau sign-in failed (${res.status})`);
  const data = await res.json();
  return { token: data.credentials.token, siteId: data.credentials.site.id };
}

async function tableauSignOut(token) {
  await fetch(`${TABLEAU_API}/auth/signout`, {
    method: 'POST',
    headers: { 'X-Tableau-Auth': token }
  }).catch(() => {});
}

async function fetchPhaseList() {
  const { patName, patSecret } = await chrome.storage.local.get(['patName', 'patSecret']);
  if (!patName || !patSecret) throw new Error('Tableau credentials not configured');
  const { token, siteId } = await tableauSignIn(patName, patSecret);
  try {
    const viewRes = await fetch(
      `${TABLEAU_API}/sites/${siteId}/views?filter=viewUrlName:eq:${TABLEAU_VIEW_URL_NAME}`,
      { headers: { 'X-Tableau-Auth': token, 'Accept': 'application/json' } }
    );
    if (!viewRes.ok) throw new Error(`Failed to query views (${viewRes.status})`);
    const viewData = await viewRes.json();
    const views = viewData.views?.view;
    if (!views || views.length === 0) throw new Error(`View "${TABLEAU_VIEW_URL_NAME}" not found`);
    const viewId = views[0].id;

    const dataRes = await fetch(
      `${TABLEAU_API}/sites/${siteId}/views/${viewId}/data`,
      { headers: { 'X-Tableau-Auth': token, 'Accept': 'text/csv' } }
    );
    if (!dataRes.ok) throw new Error(`Failed to fetch view data (${dataRes.status})`);
    const csvText = await dataRes.text();

    const wb   = XLSX.read(csvText, { type: 'string' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

    return rows
      .filter(r => r[PHASE_NAME_COL] && r[PHASE_ID_COL])
      .map(r => ({
        phase:   String(r[PHASE_NAME_COL]).trim(),
        phaseId: String(r[PHASE_ID_COL]).trim()
      }))
      .sort((a, b) => a.phase.localeCompare(b.phase));
  } finally {
    await tableauSignOut(token);
  }
}

function populatePhaseDropdown(phases) {
  const sel = document.getElementById('phaseSelect');
  sel.innerHTML = '<option value="">— select a phase —</option>';
  phases.forEach(({ phase, phaseId }) => {
    const opt = document.createElement('option');
    opt.value = phaseId;
    opt.textContent = phase;
    sel.appendChild(opt);
  });
}

async function initTableau(forceRefresh = false) {
  const statusEl = document.getElementById('connectionStatus');
  statusEl.textContent = 'Connecting…';
  statusEl.style.color = '#73b4ae';
  try {
    const phases = await fetchPhaseList();
    populatePhaseDropdown(phases);
    statusEl.textContent = `✓ ${phases.length} phases loaded`;
  } catch (err) {
    statusEl.textContent = `⚠ ${err.message}`;
    statusEl.style.color = '#fcc084';
    if (!forceRefresh) document.getElementById('credForm').style.display = 'grid';
  }
}

// ── Handing toolbar ───────────────────────────────────────────────────────────
function setupHandingToolbar() {
  const btns  = document.querySelectorAll('.handing-mode-btn');
  const mapEl = document.getElementById('map');

  function setMode(handing) {
    activeHanding = handing;
    btns.forEach(b => b.classList.toggle('active', b.dataset.handing === handing));
    mapEl.classList.remove('map-mode-ring-left', 'map-mode-ring-right');
    if (handing === 'Left')  mapEl.classList.add('map-mode-ring-left');
    if (handing === 'Right') mapEl.classList.add('map-mode-ring-right');
    if (svgEl) svgEl.style.cursor = handing ? 'crosshair' : 'grab';
  }

  btns.forEach(btn => btn.addEventListener('click', () => setMode(btn.dataset.handing)));

  document.addEventListener('keydown', e => {
    if (e.target.matches('input, select, [contenteditable]')) return;
    if (e.key === 'l' || e.key === 'L') setMode('Left');
    if (e.key === 'r' || e.key === 'R') setMode('Right');
    if (e.key === 'Escape') setMode('');
  });
}

setupHandingToolbar();

// ── Tableau connection init ───────────────────────────────────────────────────
(async () => {
  document.getElementById('toggleCredBtn').addEventListener('click', () => {
    const form = document.getElementById('credForm');
    form.style.display = form.style.display === 'none' ? 'grid' : 'none';
  });
  document.getElementById('saveCredBtn').addEventListener('click', async () => {
    const patName   = document.getElementById('patName').value.trim();
    const patSecret = document.getElementById('patSecret').value.trim();
    if (!patName || !patSecret) { alert('Enter both PAT Name and PAT Secret'); return; }
    await chrome.storage.local.set({ patName, patSecret });
    document.getElementById('credForm').style.display = 'none';
    await initTableau(true);
  });
  document.getElementById('refreshPhasesBtn').addEventListener('click', () => initTableau(true));

  const saved = await chrome.storage.local.get(['patName', 'patSecret']);
  if (saved.patName) document.getElementById('patName').value = saved.patName;

  if (saved.patName && saved.patSecret) {
    await initTableau();
  } else {
    document.getElementById('connectionStatus').textContent = 'Not configured';
    document.getElementById('credForm').style.display = 'grid';
  }
})();

// ── Main handler ──────────────────────────────────────────────────────────────
document.getElementById('processBtn').addEventListener('click', async () => {
  const phaseSelect = document.getElementById('phaseSelect');
  const phaseId     = phaseSelect.value;
  const phaseName   = phaseSelect.options[phaseSelect.selectedIndex]?.text || phaseId;
  if (!phaseId) { setStatus('error', 'Please select a Phase.'); return; }
  if (!fileInput.files[0]) { setStatus('error', 'Please select a spreadsheet file.'); return; }

  const processBtn = document.getElementById('processBtn');
  processBtn.disabled = true;
  document.getElementById('previewSection').classList.add('hidden');

  try {
    // 1. Fetch GIS data with geometry in WGS-84
    setStatus('loading', 'Fetching GIS data from ArcGIS…', true);

    const url = 'https://services7.arcgis.com/WusDoPJONiFauKEv/arcgis/rest/services/BH_Parcels_View_(Public)/FeatureServer/1/query';
    const params = new URLSearchParams({
      outFields:         'PhaseID,LotNum,BlockNum,ST_NUM,ST_NAME',
      where:             `PhaseID=${phaseId}`,
      f:                 'json',
      returnGeometry:    'true',
      outSR:             '4326',
      orderByFields:     'OBJECTID DESC',
      resultRecordCount: '2000'
    });

    const response = await fetch(`${url}?${params}`);
    if (!response.ok) throw new Error(`HTTP error — status: ${response.status}`);
    const data = await response.json();
    if (data.error) throw new Error(data.error.message);
    const features = data.features || [];

    // 2. Read spreadsheet
    setStatus('loading', `Fetched ${features.length} GIS records. Reading spreadsheet…`, true);

    const file = fileInput.files[0];
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

    if (!jsonData.length) throw new Error('No data found in the spreadsheet.');
    const firstRow = jsonData[0];
    if (!('Lot #' in firstRow))
      throw new Error('Spreadsheet must have a "Lot #" column. Found: ' + Object.keys(firstRow).join(', '));
    const hasBlockColumn = 'Block #' in firstRow;

    // 3. Blend
    setStatus('loading', 'Blending data…', true);

    const lookupMap = new Map();
    features.forEach(f => {
      const lot   = String(f.attributes.LotNum   || '').trim();
      const block = String(f.attributes.BlockNum || '').trim();
      lookupMap.set(block ? `${lot}|${block}` : lot, f);
    });

    const blended = [];
    jsonData.forEach(row => {
      const lot   = String(row['Lot #']   || '').trim();
      const block = hasBlockColumn ? String(row['Block #'] || '').trim() : '';
      const key   = block ? `${lot}|${block}` : lot;

      let match = lookupMap.get(key);
      if (!match && !block) {
        const lotN = parseInt(lot);
        if (!isNaN(lotN)) {
          for (const [k, v] of lookupMap.entries()) {
            if (!k.includes('|') && parseInt(k) === lotN) { match = v; break; }
          }
        }
      }

      if (match) {
        blended.push({
          ...row,
          PhaseID:  match.attributes.PhaseID,
          LotNum:   match.attributes.LotNum,
          BlockNum: match.attributes.BlockNum,
          ST_NUM:   match.attributes.ST_NUM,
          ST_NAME:  match.attributes.ST_NAME,
          Premium:  '',
          Handing:  ''
        });
      }
    });

    // 4. Preview panels
    const gisLots    = features.map(f => f.attributes.LotNum).sort((a, b) => parseInt(a) - parseInt(b));
    const spreadLots = jsonData.map(r => r['Lot #']).sort((a, b) => a - b);
    const matchPct   = jsonData.length ? Math.round((blended.length / jsonData.length) * 100) : 0;

    document.getElementById('gisBadge').textContent    = `${features.length} records`;
    document.getElementById('spreadBadge').textContent = `${jsonData.length} rows`;
    document.getElementById('gisMeta').textContent     = `Phase ${phaseId} · Lots ${gisLots[0]}–${gisLots[gisLots.length - 1]}`;
    document.getElementById('spreadMeta').textContent  = `${file.name} · Lots ${Math.min(...spreadLots)}–${Math.max(...spreadLots)}`;

    renderPreviewTable(features.slice(0, 10).map(f => f.attributes), 'gisTable');
    renderPreviewTable(jsonData.slice(0, 10), 'spreadTable');

    document.getElementById('matchCount').textContent = blended.length;
    document.getElementById('matchLabel').textContent = `of ${jsonData.length} rows matched (${matchPct}%)`;

    document.getElementById('previewSection').classList.remove('hidden');

    // 5. SVG map + editable table
    initMapAndEditor(features, blended);

    setStatus(
      blended.length ? 'success' : 'error',
      blended.length
        ? `${blended.length} rows blended. Click a parcel or row to edit, then download.`
        : 'No matches found — Lot numbers may not overlap between GIS data and the spreadsheet.'
    );

    // 6. Download — reads live DOM so edits are captured
    document.getElementById('downloadBtn').onclick = () => {
      const finalData = getEditedData();
      const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(finalData));
      chrome.downloads.download({
        url:      'data:text/csv;charset=utf-8,' + encodeURIComponent(csv),
        filename: `blended_Phase${phaseId}.csv`
      });
      setStatus('success', 'Download started!');
    };

  } catch (error) {
    setStatus('error', 'Error: ' + error.message);
  } finally {
    processBtn.disabled = false;
  }
});
