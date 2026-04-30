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

const GIS_COLS = new Set(['House Number', 'Street', 'Zip']);

// ── Tableau connection ────────────────────────────────────────────────────────
const TABLEAU_SERVER        = 'https://10ay.online.tableau.com';
const TABLEAU_SITE          = 'betenboughhomes';
const TABLEAU_API           = `${TABLEAU_SERVER}/api/3.22`;
const PHASE_NAME_COL        = 'Phase';
const PHASE_ID_COL          = 'Phase ID';
const TABLEAU_VIEW_URL_NAME = 'PhaseList';

// ── Classification constants ──────────────────────────────────────────────────
const CLASSIFICATION_OPTIONS = [
  '40 Foot Compact', '50 Foot Compact', '45 Foot', '50 Foot',
  '50 Foot RSV', '60 Foot', '75 Foot', '85 Foot'
];
const CLASSIFICATION_COLORS = {
  '40 Foot Compact': '#fcc084',
  '50 Foot Compact': '#db5d00',
  '45 Foot':         '#ccd47e',
  '50 Foot':         '#606426',
  '50 Foot RSV':     '#73b4ae',
  '60 Foot':         '#375957',
  '75 Foot':         '#243a38',
  '85 Foot':         '#1a2725'
};

function rowKey(attrs) {
  const lot   = String(attrs.LotNum   || attrs['Lot #']   || '').trim();
  const block = String(attrs.BlockNum || attrs['Block #'] || '').trim();
  return block ? `${lot}|${block}` : lot;
}

// ── SVG map state ─────────────────────────────────────────────────────────────
let svgEl           = null;
let svgActiveGroup  = null;
let svgMatchedKeys  = new Set();
let svgContentGroup = null;
let svgTransform    = { x: 0, y: 0, s: 1 };
const LABEL_ZOOM_THRESHOLD = 1.5; // address labels appear at 2× initial zoom

// Handing assignment state
let handingMap        = new Map();   // key → 'Left' | 'Right' | ''
let activeHanding     = '';          // currently active paint mode
let activeToolset     = 'handing';   // 'handing' | 'address' | 'classification'
let classificationMap = new Map();   // key → classification string
let currentFeatures   = [];          // stored for toolset re-render
let addressLabelColor = localStorage.getItem('addressLabelColor') || '#dc2626';

function updateZoomHint() {
  const hint = document.getElementById('zoomHint');
  if (!hint) return;
  hint.style.display = (activeToolset === 'address' && svgTransform.s < LABEL_ZOOM_THRESHOLD) ? 'block' : 'none';
}

function applyMapTransform() {
  if (!svgContentGroup) return;
  const { x, y, s } = svgTransform;
  svgContentGroup.setAttribute('transform', `translate(${x},${y}) scale(${s})`);
  if (svgEl && activeToolset === 'address') {
    const show  = s >= LABEL_ZOOM_THRESHOLD;
    // Counterscale so screen size stays constant (11px / 10px) regardless of zoom level
    const size1 = (12.0 * LABEL_ZOOM_THRESHOLD / s).toFixed(2);
    const size2 = (8.0 * LABEL_ZOOM_THRESHOLD / s).toFixed(2);
    svgEl.querySelectorAll('text.address-label').forEach(txt => {
      txt.setAttribute('opacity', show ? '1' : '0');
      const spans = txt.querySelectorAll('tspan');
      if (spans[0]) spans[0].setAttribute('font-size', size1);
      if (spans[1]) spans[1].setAttribute('font-size', size2);
    });
  }
  updateZoomHint();
}

function styleFeatureGroup(fg, state) {
  const key     = fg.dataset.key;
  const matched = svgMatchedKeys.has(key);

  fg.querySelectorAll('path').forEach(p => {
    if (!matched) {
      p.setAttribute('fill', '#e4e2bd'); p.setAttribute('fill-opacity', '0.35');
      p.setAttribute('stroke', '#a0a87a'); p.setAttribute('stroke-width', '1');
      return;
    }

    let fill, stroke;
    if (activeToolset === 'handing') {
      const h = handingMap.get(key) || '';
      fill   = h === 'Left' ? '#ede9fe' : h === 'Right' ? '#fee2e2' : '#73b4ae';
      stroke = h === 'Left' ? '#7c3aed' : h === 'Right' ? '#dc2626' : '#375957';
    } else if (activeToolset === 'classification') {
      const cls = classificationMap.get(key) || '';
      fill   = CLASSIFICATION_COLORS[cls] || '#73b4ae';
      stroke = CLASSIFICATION_COLORS[cls] || '#375957';
    } else {
      fill = '#73b4ae'; stroke = '#375957';
    }

    if (state === 'selected') {
      const h = activeToolset === 'handing' ? (handingMap.get(key) || '') : '';
      p.setAttribute('fill',   h === 'Left' ? '#ddd6fe' : h === 'Right' ? '#fecaca' : '#db5d00');
      p.setAttribute('fill-opacity', '0.82');
      p.setAttribute('stroke', h === 'Left' ? '#5b21b6' : h === 'Right' ? '#b91c1c' : '#a34500');
      p.setAttribute('stroke-width', '2.5');
    } else if (state === 'hover') {
      p.setAttribute('fill', fill); p.setAttribute('fill-opacity', '0.78');
      p.setAttribute('stroke', stroke); p.setAttribute('stroke-width', '1.8');
    } else {
      p.setAttribute('fill', fill);
      p.setAttribute('fill-opacity', (activeToolset === 'handing' && handingMap.get(key)) || activeToolset === 'classification' ? '0.68' : '0.5');
      p.setAttribute('stroke', stroke); p.setAttribute('stroke-width', '1.2');
    }
  });

  // Text label color — only relevant for handing toolset
  if (activeToolset === 'handing') {
    const h = handingMap.get(key) || '';
    fg.querySelectorAll('text').forEach(t =>
      t.setAttribute('fill', h === 'Left' ? '#5b21b6' : h === 'Right' ? '#991b1b' : '#1a2725')
    );
  }
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
  tr.classList.remove('handing-left', 'handing-right');
  if (handing === 'Left')  tr.classList.add('handing-left');
  if (handing === 'Right') tr.classList.add('handing-right');
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

  currentFeatures = features;

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

    // Centroid label / marker (toolset-aware)
    const ring0 = f.geometry.rings[0];
    if (ring0?.length) {
      const cx = ring0.reduce((s, p) => s + p[0], 0) / ring0.length;
      const cy = ring0.reduce((s, p) => s + p[1], 0) / ring0.length;
      const [sx, sy] = project([cx, cy]);

      if (activeToolset === 'address') {
        const stNum  = String(f.attributes.ST_NUM  || '');
        const stName = String(f.attributes.ST_NAME || '');
        if (stNum || stName) {
          const txt = document.createElementNS(NS, 'text');
          txt.setAttribute('class', 'address-label');
          txt.setAttribute('x', sx); txt.setAttribute('y', sy);
          txt.setAttribute('text-anchor', 'middle');
          txt.setAttribute('dominant-baseline', 'middle');
          txt.setAttribute('font-family', '-apple-system, BlinkMacSystemFont, sans-serif');
          txt.setAttribute('fill', addressLabelColor);
          txt.setAttribute('stroke', addressLabelColor === '#ffffff' ? 'rgba(0,0,0,0.65)' : 'rgba(255,255,255,0.9)');
          txt.setAttribute('stroke-width', '0.5');
          txt.setAttribute('paint-order', 'stroke');
          txt.setAttribute('pointer-events', 'none');
          txt.setAttribute('opacity', svgTransform.s >= LABEL_ZOOM_THRESHOLD ? '1' : '0');
          const initSize1 = (12.0 * LABEL_ZOOM_THRESHOLD / svgTransform.s).toFixed(2);
          const initSize2 = (8.0 * LABEL_ZOOM_THRESHOLD / svgTransform.s).toFixed(2);
          const span1 = document.createElementNS(NS, 'tspan');
          span1.setAttribute('x', sx); span1.setAttribute('dy', '-3'); span1.setAttribute('font-size', initSize1);
          span1.textContent = stNum;
          const span2 = document.createElementNS(NS, 'tspan');
          span2.setAttribute('x', sx); span2.setAttribute('dy', '7'); span2.setAttribute('font-size', initSize2);
          span2.textContent = stName;
          txt.appendChild(span1); txt.appendChild(span2);
          fg.appendChild(txt);
        }
      } else {
        const label = activeToolset === 'classification' ? '' : String(f.attributes.LotNum || '');
        if (label) {
          const txt = document.createElementNS(NS, 'text');
          txt.setAttribute('x', sx); txt.setAttribute('y', sy);
          txt.setAttribute('text-anchor', 'middle'); txt.setAttribute('dominant-baseline', 'middle');
          txt.setAttribute('font-size', '8');
          txt.setAttribute('font-family', '-apple-system, BlinkMacSystemFont, sans-serif');
          txt.setAttribute('fill', '#1a2725'); txt.setAttribute('pointer-events', 'none');
          txt.textContent = label;
          fg.appendChild(txt);
        }
      }
    }

    // Hover
    fg.addEventListener('mouseenter', () => {
      if (fg !== svgActiveGroup) styleFeatureGroup(fg, 'hover');
    });
    fg.addEventListener('mouseleave', () => {
      if (fg !== svgActiveGroup) styleFeatureGroup(fg, 'default');
    });

    // Click — toolset-aware
    if (matched) {
      fg.addEventListener('click', e => {
        if (dragged) return;
        e.stopPropagation();
        if (activeToolset === 'handing' && activeHanding !== '') {
          handingMap.set(key, activeHanding);
          styleFeatureGroup(fg, fg === svgActiveGroup ? 'selected' : 'default');
          syncTableHanding(key, activeHanding);
        } else {
          if (svgActiveGroup && svgActiveGroup !== fg) styleFeatureGroup(svgActiveGroup, 'default');
          svgActiveGroup = fg;
          styleFeatureGroup(fg, 'selected');
          if (activeToolset === 'address')        showAddressPopover(key, e);
          if (activeToolset === 'classification') showClassificationPopover(key, e);
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
  svg.addEventListener('click', () => {
    // Click on map background (feature clicks stopPropagation, so this only fires for blank space)
    if (svgActiveGroup) { styleFeatureGroup(svgActiveGroup, 'default'); svgActiveGroup = null; }
    document.querySelectorAll('#editTableWrap tr.row-selected').forEach(r => r.classList.remove('row-selected'));
  });

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
                return `<td class="col-gis" data-col="${k}" title="${row[k] ?? ''}">${row[k] ?? ''}</td>`;
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
              } else if (k === 'Homesite Classification') {
                return `<td class="col-handing">
                  <select class="handing-select" data-col="Homesite Classification">
                    <option value="" ${!row[k] ? 'selected' : ''}>—</option>
                    ${CLASSIFICATION_OPTIONS.map(opt =>
                      `<option value="${opt}" ${row[k] === opt ? 'selected' : ''}>${opt}</option>`
                    ).join('')}
                  </select>
                </td>`;
              } else if (k === 'Plat Name') {
                return `<td class="col-platname" data-col="Plat Name" title="${row[k] ?? ''}">${row[k] ?? ''}</td>`;
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

  // Handing — mark cell edited + sync map color + row highlight
  wrap.querySelectorAll('select[data-col="Handing"]').forEach(sel => {
    sel.addEventListener('change', () => {
      const key = sel.closest('tr').dataset.key;
      const tr  = sel.closest('tr');
      handingMap.set(key, sel.value);
      tr.classList.remove('handing-left', 'handing-right');
      if (sel.value === 'Left')  tr.classList.add('handing-left');
      if (sel.value === 'Right') tr.classList.add('handing-right');
      updateMapFeatureStyle(key);
      sel.closest('td').classList.toggle('cell-edited', sel.value !== '');
    });
  });

  // Classification — sync map color
  wrap.querySelectorAll('select[data-col="Homesite Classification"]').forEach(sel => {
    sel.addEventListener('change', () => {
      const key = sel.closest('tr').dataset.key;
      classificationMap.set(key, sel.value);
      const idx = blendedData.findIndex(r => rowKey(r) === key);
      if (idx !== -1) blendedData[idx]['Homesite Classification'] = sel.value;
      updateMapFeatureStyle(key);
      sel.closest('td').classList.toggle('cell-edited', !!sel.value);
    });
  });

  // Apply initial handing row highlights for data loaded from spreadsheet
  wrap.querySelectorAll('tr[data-key]').forEach(tr => {
    const sel = tr.querySelector('select[data-col="Handing"]');
    if (sel?.value === 'Left')  tr.classList.add('handing-left');
    if (sel?.value === 'Right') tr.classList.add('handing-right');
  });

  // Click a GIS cell / row → select on map
  wrap.querySelectorAll('tr[data-key]').forEach(tr => {
    tr.addEventListener('click', e => {
      if (e.target.classList.contains('col-sheet')) return;
      document.querySelectorAll('#editTableWrap tr.row-selected').forEach(r => r.classList.remove('row-selected'));
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
      edited[input.dataset.col] = input.value === '' ? 0 : parseInt(input.value, 10);
    });
    // Handing — store as string
    tr.querySelectorAll('select[data-col]').forEach(sel => {
      edited[sel.dataset.col] = sel.value;
    });
    return edited;
  });
}

function initMapAndEditor(features, blended) {
  blendedData       = blended;
  currentFeatures   = features;
  handingMap = new Map(
    blended
      .filter(r => r['Handing'])
      .map(r => [rowKey(r), r['Handing']])
  );
  classificationMap = new Map(
    blended
      .filter(r => r['Homesite Classification'])
      .map(r => [rowKey(r), r['Homesite Classification']])
  );
  activeToolset     = 'handing';
  activeHanding     = '';
  document.querySelectorAll('.toolset-tab').forEach(t =>
    t.classList.toggle('active', t.dataset.toolset === 'handing')
  );
  renderToolbarContent('handing');
  renderLegendContent('handing');
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
      { headers: { 'X-Tableau-Auth': token } }
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

let allPhases = [];

function renderPhaseOptions(filter) {
  const query = filter.toLowerCase();
  const list  = query ? allPhases.filter(p => p.phase.toLowerCase().includes(query)) : allPhases;

  // Sync hidden select (processBtn reads .value and .options[].text)
  const sel  = document.getElementById('phaseSelect');
  const prev = sel.value;
  sel.innerHTML = '<option value="">—</option>';
  list.forEach(({ phase, phaseId }) => {
    const opt = document.createElement('option');
    opt.value = phaseId; opt.textContent = phase; sel.appendChild(opt);
  });
  if (prev) sel.value = prev;

  // Build visible list
  const container = document.getElementById('phasePickerOptions');
  if (!container) return;
  if (!list.length) {
    container.innerHTML = '<div class="phase-option-empty">No phases match</div>'; return;
  }
  container.innerHTML = '';
  list.forEach(({ phase, phaseId }) => {
    const div = document.createElement('div');
    div.className = 'phase-option' + (sel.value === phaseId ? ' selected' : '');
    div.textContent = phase; div.dataset.phaseId = phaseId;
    div.addEventListener('mousedown', e => { e.preventDefault(); selectPhase(phaseId, phase); });
    container.appendChild(div);
  });
}

function selectPhase(phaseId, phaseName) {
  document.getElementById('phaseSelect').value = phaseId;
  document.getElementById('phasePickerLabel').textContent = phaseName || '— select a phase —';
  closePickerPanel();
  document.getElementById('phaseSearch').value = '';
  renderPhaseOptions('');
}

function openPickerPanel() {
  document.getElementById('phasePickerPanel').style.display = 'block';
  document.getElementById('phasePickerTrigger').classList.add('open');
  document.getElementById('phaseSearch').focus();
}

function closePickerPanel() {
  document.getElementById('phasePickerPanel').style.display = 'none';
  document.getElementById('phasePickerTrigger').classList.remove('open');
}

function populatePhaseDropdown(phases) {
  allPhases = phases;
  const searchEl = document.getElementById('phaseSearch');
  if (searchEl) searchEl.value = '';
  renderPhaseOptions('');
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

// ── Toolset switcher ──────────────────────────────────────────────────────────
function renderToolbarContent(toolset) {
  const tb = document.getElementById('mapToolbar');
  if (!tb) return;
  if (toolset === 'handing') {
    tb.innerHTML = `
      <span class="toolbar-label">Handing</span>
      <div class="handing-btns">
        <button class="handing-mode-btn active" data-handing="">Navigate</button>
        <button class="handing-mode-btn mode-left"  data-handing="Left">&#9664; Left</button>
        <button class="handing-mode-btn mode-right" data-handing="Right">Right &#9654;</button>
      </div>
      <span class="toolbar-hint">L &nbsp;&middot;&nbsp; R &nbsp;&middot;&nbsp; Esc</span>`;
    attachHandingToolbarListeners();
  } else if (toolset === 'address') {
    const colors = [
      { hex: '#dc2626', title: 'Red'   },
      { hex: '#1d4ed8', title: 'Blue'  },
      { hex: '#ffffff', title: 'White' },
      { hex: '#1a2725', title: 'Black' }
    ];
    tb.innerHTML = `
      <span class="toolbar-label">Address Editor</span>
      <div class="addr-color-btns">
        ${colors.map(c => `<button class="addr-color-btn${addressLabelColor === c.hex ? ' active' : ''}"
          data-color="${c.hex}" style="background:${c.hex};" title="${c.title}"></button>`).join('')}
      </div>
      <span class="toolbar-hint" style="margin-left:6px">Click a parcel to edit address</span>`;
    attachAddressColorListeners();
  } else {
    tb.innerHTML = `
      <span class="toolbar-label">Classification</span>
      <span class="toolbar-hint" style="margin-left:0">Click a parcel to assign a homesite classification</span>`;
  }
}

function renderLegendContent(toolset) {
  const lg = document.getElementById('mapLegend');
  if (!lg) return;
  if (toolset === 'handing') {
    lg.innerHTML = `
      <div class="legend-item"><div class="legend-swatch swatch-unassigned"></div> Unassigned</div>
      <div class="legend-item"><div class="legend-swatch swatch-left"></div> Left</div>
      <div class="legend-item"><div class="legend-swatch swatch-right"></div> Right</div>
      <div class="legend-item"><div class="legend-swatch swatch-unmatched"></div> No match</div>`;
  } else if (toolset === 'address') {
    lg.innerHTML = `
      <div class="legend-item"><div class="legend-swatch swatch-unassigned"></div> Matched</div>
      <div class="legend-item"><div class="legend-swatch swatch-unmatched"></div> No match</div>`;
  } else {
    const swatches = CLASSIFICATION_OPTIONS.map(opt =>
      `<div class="legend-item">
        <div class="legend-swatch" style="background:${CLASSIFICATION_COLORS[opt]};border:1.5px solid ${CLASSIFICATION_COLORS[opt]};"></div>
        ${opt}
      </div>`
    ).join('');
    lg.innerHTML = swatches +
      `<div class="legend-item"><div class="legend-swatch" style="background:#c0dae1;border:1.5px dashed #73b4ae;"></div> Unassigned</div>
       <div class="legend-item"><div class="legend-swatch swatch-unmatched"></div> No match</div>`;
  }
}

// ── Map popovers ──────────────────────────────────────────────────────────────
function showMapPopover(html, e) {
  const popover = document.getElementById('mapPopover');
  const panel   = document.querySelector('.map-panel');
  const rect    = panel.getBoundingClientRect();
  popover.innerHTML = html;
  popover.style.display = 'block';
  let left = e.clientX - rect.left + 12;
  let top  = e.clientY - rect.top  + 12;
  popover.style.left = left + 'px';
  popover.style.top  = top  + 'px';
  requestAnimationFrame(() => {
    const pr   = popover.getBoundingClientRect();
    const panR = panel.getBoundingClientRect();
    if (pr.right  > panR.right  - 8) left -= pr.width  + 24;
    if (pr.bottom > panR.bottom - 8) top  -= pr.height + 24;
    popover.style.left = Math.max(8, left) + 'px';
    popover.style.top  = Math.max(8, top)  + 'px';
  });
}

function hideMapPopover() {
  const p = document.getElementById('mapPopover');
  if (p) p.style.display = 'none';
}

function showAddressPopover(key, e) {
  const row = blendedData.find(r => rowKey(r) === key);
  if (!row) return;
  showMapPopover(`
    <div class="popover-title">Edit Address</div>
    <div class="popover-field">
      <span class="popover-label">House Number</span>
      <input class="popover-input" id="popStNum" value="${String(row['House Number'] ?? '').replace(/"/g, '&quot;')}">
    </div>
    <div class="popover-field">
      <span class="popover-label">Street</span>
      <input class="popover-input" id="popStName" value="${String(row['Street'] ?? '').replace(/"/g, '&quot;')}">
    </div>
    <div class="popover-actions">
      <button class="popover-cancel" id="popCancel">Cancel</button>
      <button class="popover-save"   id="popSave">Save</button>
    </div>`, e);

  document.getElementById('popSave').onclick = () => {
    const stNum  = document.getElementById('popStNum').value.trim();
    const stName = document.getElementById('popStName').value.trim();
    const idx = blendedData.findIndex(r => rowKey(r) === key);
    if (idx !== -1) { blendedData[idx]['House Number'] = stNum; blendedData[idx]['Street'] = stName; }
    const wrap = document.getElementById('editTableWrap');
    const tr   = wrap?.querySelector(`tr[data-key="${key}"]`);
    if (tr) {
      tr.querySelectorAll('td.col-gis[data-col]').forEach(td => {
        if (td.dataset.col === 'House Number') td.textContent = stNum;
        if (td.dataset.col === 'Street')       td.textContent = stName;
      });
    }
    if (svgEl) {
      const fg  = svgEl.querySelector(`g[data-key="${key}"]`);
      const txt = fg?.querySelector('text');
      if (txt) txt.textContent = [stNum, stName].filter(Boolean).join(' ');
    }
    hideMapPopover();
  };
  document.getElementById('popCancel').onclick = hideMapPopover;
}

function showClassificationPopover(key, e) {
  const row     = blendedData.find(r => rowKey(r) === key);
  const current = classificationMap.get(key) || row?.['Homesite Classification'] || '';
  const options = CLASSIFICATION_OPTIONS.map(opt =>
    `<option value="${opt}" ${current === opt ? 'selected' : ''}>${opt}</option>`
  ).join('');
  showMapPopover(`
    <div class="popover-title">Homesite Classification</div>
    <div class="popover-field">
      <select class="popover-select" id="popClassSel">
        <option value="">— unassigned —</option>
        ${options}
      </select>
    </div>
    <div class="popover-actions">
      <button class="popover-cancel" id="popCancel">Cancel</button>
      <button class="popover-save"   id="popSave">Save</button>
    </div>`, e);

  document.getElementById('popSave').onclick = () => {
    const cls = document.getElementById('popClassSel').value;
    classificationMap.set(key, cls);
    const idx = blendedData.findIndex(r => rowKey(r) === key);
    if (idx !== -1) blendedData[idx]['Homesite Classification'] = cls;
    const wrap = document.getElementById('editTableWrap');
    const sel  = wrap?.querySelector(`tr[data-key="${key}"] select[data-col="Homesite Classification"]`);
    if (sel) { sel.value = cls; sel.closest('td').classList.toggle('cell-edited', !!cls); }
    updateMapFeatureStyle(key);
    // Also update the centroid rect color
    if (svgEl) {
      const fg   = svgEl.querySelector(`g[data-key="${key}"]`);
      const rect = fg?.querySelector('rect');
      if (rect) rect.setAttribute('fill', CLASSIFICATION_COLORS[cls] || '#c0dae1');
    }
    hideMapPopover();
  };
  document.getElementById('popCancel').onclick = hideMapPopover;
}

function attachAddressColorListeners() {
  document.querySelectorAll('.addr-color-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      addressLabelColor = btn.dataset.color;
      localStorage.setItem('addressLabelColor', addressLabelColor);
      document.querySelectorAll('.addr-color-btn').forEach(b =>
        b.classList.toggle('active', b.dataset.color === addressLabelColor)
      );
      if (svgEl) {
        const halo = addressLabelColor === '#ffffff' ? 'rgba(0,0,0,0.65)' : 'rgba(255,255,255,0.9)';
        svgEl.querySelectorAll('text.address-label').forEach(txt => {
          txt.setAttribute('fill', addressLabelColor);
          txt.setAttribute('stroke', halo);
        });
      }
    });
  });
}

function attachHandingToolbarListeners() {
  document.querySelectorAll('.handing-mode-btn').forEach(btn =>
    btn.addEventListener('click', () => setHandingMode(btn.dataset.handing))
  );
}

function setHandingMode(handing) {
  activeHanding = handing;
  document.querySelectorAll('.handing-mode-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.handing === handing)
  );
  const mapEl = document.getElementById('map');
  mapEl.classList.remove('map-mode-ring-left', 'map-mode-ring-right');
  if (handing === 'Left')  mapEl.classList.add('map-mode-ring-left');
  if (handing === 'Right') mapEl.classList.add('map-mode-ring-right');
  if (svgEl) svgEl.style.cursor = handing ? 'crosshair' : 'grab';
}

function switchToolset(toolset) {
  activeToolset = toolset;
  activeHanding = '';
  document.querySelectorAll('.toolset-tab').forEach(t =>
    t.classList.toggle('active', t.dataset.toolset === toolset)
  );
  document.getElementById('map').classList.remove('map-mode-ring-left', 'map-mode-ring-right');
  hideMapPopover();
  renderToolbarContent(toolset);
  renderLegendContent(toolset);
  if (currentFeatures.length) renderSVGMap(currentFeatures, blendedData);
  updateZoomHint();
}

// ── Handing toolbar ───────────────────────────────────────────────────────────
function setupHandingToolbar() {
  document.addEventListener('keydown', e => {
    if (e.target.matches('input, select, [contenteditable]')) return;
    if (e.key === 'Escape') { hideMapPopover(); if (activeToolset === 'handing') setHandingMode(''); return; }
    if (activeToolset !== 'handing') return;
    if (e.key === 'l' || e.key === 'L') setHandingMode('Left');
    if (e.key === 'r' || e.key === 'R') setHandingMode('Right');
  });
}

setupHandingToolbar();

document.querySelectorAll('.toolset-tab').forEach(tab =>
  tab.addEventListener('click', () => switchToolset(tab.dataset.toolset))
);

document.addEventListener('click', e => {
  const popover = document.getElementById('mapPopover');
  if (popover && popover.style.display !== 'none' &&
      !popover.contains(e.target) && !e.target.closest('#map')) {
    hideMapPopover();
  }
});

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
  const phasePickerTrigger = document.getElementById('phasePickerTrigger');
  const phaseSearchEl      = document.getElementById('phaseSearch');

  phasePickerTrigger.addEventListener('click', () => {
    document.getElementById('phasePickerPanel').style.display !== 'none'
      ? closePickerPanel() : openPickerPanel();
  });
  phasePickerTrigger.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ' || e.key === 'ArrowDown')
      { e.preventDefault(); openPickerPanel(); }
  });
  phaseSearchEl.addEventListener('input', e => renderPhaseOptions(e.target.value));
  phaseSearchEl.addEventListener('keydown', e => {
    if (e.key === 'Escape') closePickerPanel();
  });
  document.addEventListener('click', e => {
    if (!e.target.closest('#phasePicker')) closePickerPanel();
  });

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
      outFields:         'PhaseID,LotNum,BlockNum,ST_NUM,ST_NAME,ZIP',
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
          'House Number':            match.attributes.ST_NUM  ?? '',
          'Street':                  match.attributes.ST_NAME ?? '',
          'Plat Name':               document.getElementById('platNameInput')?.value || row['Plat Name'] || '',
          'Lot #':                   row['Lot #']         ?? '',
          'Block #':                 row['Block #']       ?? '',
          'Zip':                     match.attributes.ZIP ?? '',
          'Homesite Classification': row['Homesite Classification'] ?? '',
          'Premium':                 row['Premium'] != null && row['Premium'] !== '' ? Number(row['Premium']) : 0,
          'Handing':                 row['Handing'] ?? '',
          'Depth':                   row['Depth']         ?? '',
          'Width':                   row['Width']         ?? '',
          'Total Size':              row['Total Size']    ?? '',
          'Max Box Width':           row['Max Box Width'] ?? '',
          'Max Box Depth':           row['Max Box Depth'] ?? ''
        });
      }
    });

    // Seed Plat Name input from spreadsheet if the input is currently empty
    const platInput = document.getElementById('platNameInput');
    if (platInput && !platInput.value) {
      const platFromSpread = blended.find(r => r['Plat Name'])?.['Plat Name'];
      if (platFromSpread) platInput.value = platFromSpread;
    }

    // 4. Preview panels
    const gisLots    = features.map(f => f.attributes.LotNum).sort((a, b) => parseInt(a) - parseInt(b));
    const spreadLots = jsonData.map(r => r['Lot #']).sort((a, b) => a - b);
    const matchPct   = jsonData.length ? Math.round((blended.length / jsonData.length) * 100) : 0;

    document.getElementById('gisBadge').textContent    = `${features.length} records`;
    document.getElementById('spreadBadge').textContent = `${jsonData.length} rows`;
    document.getElementById('gisMeta').textContent     = `${phaseName} · Lots ${gisLots[0]}–${gisLots[gisLots.length - 1]}`;
    document.getElementById('spreadMeta').textContent  = `${file.name} · Lots ${Math.min(...spreadLots)}–${Math.max(...spreadLots)}`;

    renderPreviewTable(features.slice(0, 10).map(f => f.attributes), 'gisTable');
    renderPreviewTable(jsonData.slice(0, 10), 'spreadTable');

    document.getElementById('matchCount').textContent = blended.length;
    document.getElementById('matchLabel').textContent = `of ${jsonData.length} rows matched (${matchPct}%)`;

    document.getElementById('previewSection').classList.remove('hidden');

    // 5. SVG map + editable table
    initMapAndEditor(features, blended);

    // PlatName — single input fills all rows
    const platNameInput = document.getElementById('platNameInput');
    platNameInput.oninput = () => {
      const val = platNameInput.value;
      blendedData.forEach(r => { r['Plat Name'] = val; });
      document.querySelectorAll('#editTableWrap td[data-col="Plat Name"]').forEach(td => {
        td.textContent = val;
        td.title       = val;
      });
    };

    setStatus(
      blended.length ? 'success' : 'error',
      blended.length
        ? `${blended.length} rows blended. Click a parcel or row to edit, then download.`
        : 'No matches found — Lot numbers may not overlap between GIS data and the spreadsheet.'
    );

    // 6. Download — reads live DOM so edits are captured
    document.getElementById('downloadBtn').onclick = () => {
      const finalData  = getEditedData();
      const sheetName  = `${phaseName || 'Phase' + phaseId}_ImPortal`.slice(0, 31); // Excel sheet name max 31 chars
      const wb         = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(finalData), sheetName);
      const xlsxBytes  = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
      chrome.downloads.download({
        url:      'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + xlsxBytes,
        filename: `${sheetName}.xlsx`
      });
      setStatus('success', 'Download started!');
    };

  } catch (error) {
    setStatus('error', 'Error: ' + error.message);
  } finally {
    processBtn.disabled = false;
  }
});
