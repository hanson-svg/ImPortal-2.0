document.getElementById('processBtn').addEventListener('click', async () => {
  console.log('Button clicked');
  const phaseId = document.getElementById('phaseId').value.trim();
  if (!phaseId) {
    alert('Please enter a Phase ID');
    return;
  }

  const fileInput = document.getElementById('fileInput');
  if (!fileInput.files[0]) {
    alert('Please select a file');
    return;
  }

  const status = document.getElementById('status');
  status.textContent = 'Fetching GIS data...';
  console.log('Starting process for Phase ID:', phaseId);

  try {
    // Fetch GIS data
    const url = 'https://services7.arcgis.com/WusDoPJONiFauKEv/arcgis/rest/services/BH_Parcels_View_(Public)/FeatureServer/1/query';
    const params = new URLSearchParams({
      outFields: 'PhaseID,LotNum,BlockNum,ST_NUM,ST_NAME',
      where: `PhaseID=${phaseId}`,
      f: 'json',
      returnGeometry: 'false',
      orderByFields: 'OBJECTID DESC',
      resultRecordCount: '2000'
    });
    console.log('Fetching:', `${url}?${params}`);
    const response = await fetch(`${url}?${params}`);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    if (data.error) {
      throw new Error(data.error.message);
    }
    const features = data.features || [];
    console.log(`Fetched ${features.length} features`);

    status.textContent = `Fetched ${features.length} records for Phase ${phaseId}`;

    // Read spreadsheet
    const file = fileInput.files[0];
    console.log('Reading file:', file.name);
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet);
    console.log(`Parsed ${jsonData.length} rows`);

    if (jsonData.length === 0) {
      throw new Error('No data in spreadsheet');
    }

    const firstRow = jsonData[0];
    console.log('=== SPREADSHEET ANALYSIS ===');
    console.log('All column names:', Object.keys(firstRow));
    console.log('First 3 rows as objects:');
    jsonData.slice(0, 3).forEach((row, idx) => {
      console.log(`Row ${idx}:`, JSON.stringify(row, null, 2));
    });
    
    if (!('Lot #' in firstRow)) {
      throw new Error('Spreadsheet must have "Lot #" column. Found: ' + Object.keys(firstRow).join(', '));
    }
    const hasBlockColumn = 'Block #' in firstRow;
    if (!hasBlockColumn) {
      console.log('Note: "Block #" column not found, will match on Lot # only');
    }

    status.textContent = `Parsed ${jsonData.length} rows from spreadsheet`;

    // Blend data - build lookup map for order-independent matching
    const lookupMap = new Map();
    console.log('=== GIS DATA ANALYSIS ===');
    console.log('First 3 GIS records as full objects:');
    features.slice(0, 3).forEach((f, idx) => {
      console.log(`GIS ${idx}:`, JSON.stringify(f.attributes, null, 2));
    });
    
    features.forEach(f => {
      const lotNum = String(f.attributes.LotNum || '').trim();
      const blockNum = String(f.attributes.BlockNum || '').trim();
      const key = blockNum ? `${lotNum}|${blockNum}` : lotNum;
      lookupMap.set(key, f);
      if (lookupMap.size <= 10) {
        console.log(`GIS [${lookupMap.size}]: key="${key}", LotNum="${f.attributes.LotNum}", BlockNum="${f.attributes.BlockNum}"`);
      }
    });

    console.log(`Built lookup map with ${lookupMap.size} entries`);
    console.log('=== BLEND ANALYSIS ===');
    console.log('Lookup keys:', Array.from(lookupMap.keys()).slice(0, 20));

    const blended = [];
    jsonData.forEach((row, idx) => {
      const lotNum = String(row['Lot #'] || '').trim();
      const blockNum = hasBlockColumn ? String(row['Block #'] || '').trim() : '';
      const key = blockNum ? `${lotNum}|${blockNum}` : lotNum;
      
      let match = lookupMap.get(key);
      
      if (idx < 10) {
        console.log(`Row ${idx}: fullObject=${JSON.stringify(row)}, Lot#="${row['Lot #']}" -> lotNum="${lotNum}", key="${key}", match=${match ? 'FOUND' : 'NOT FOUND'}`);
      }
      
      if (!match && !blockNum) {
        // Try numeric matching if string didn't work
        const lotNumeric = parseInt(lotNum);
        if (!isNaN(lotNumeric)) {
          for (const [k, v] of lookupMap.entries()) {
            if (!k.includes('|')) {
              const gisNumeric = parseInt(k);
              if (!isNaN(gisNumeric) && gisNumeric === lotNumeric) {
                match = v;
                if (idx < 3) console.log(`Row ${idx}: Found numeric match: ${lotNumeric} = ${gisNumeric}`);
                break;
              }
            }
          }
        }
      }

      if (match) {
        blended.push({
          ...row,
          PhaseID: match.attributes.PhaseID,
          LotNum: match.attributes.LotNum,
          BlockNum: match.attributes.BlockNum,
          ST_NUM: match.attributes.ST_NUM,
          ST_NAME: match.attributes.ST_NAME
        });
      }
    });
    
    console.log(`Blended ${blended.length} rows`);
    console.log('First 5 GIS Lot numbers:', features.slice(0, 5).map(f => f.attributes.LotNum));
    console.log('First 5 spreadsheet Lot numbers:', jsonData.slice(0, 5).map(r => r['Lot #']));

    // Show preview
    const preview = document.getElementById('preview');
    const gisPreview = document.getElementById('gisPreview');
    const spreadPreview = document.getElementById('spreadPreview');
    const blendedPreview = document.getElementById('blendedPreview');
    
    const gisLots = features.map(f => f.attributes.LotNum).sort((a, b) => parseInt(a) - parseInt(b));
    const spreadLots = jsonData.map(r => r['Lot #']).sort((a, b) => a - b);
    
    gisPreview.textContent = `GIS Lot Range: ${gisLots[0]} - ${gisLots[gisLots.length-1]} (${gisLots.length} total)\n\nFirst 10: ${gisLots.slice(0, 10).join(', ')}\n\nFirst 5 records:\n${JSON.stringify(features.slice(0, 5).map(f => f.attributes), null, 2)}`;
    spreadPreview.textContent = `Spreadsheet Lot Range: ${Math.min(...spreadLots)} - ${Math.max(...spreadLots)} (${spreadLots.length} total)\n\nFirst 10: ${spreadLots.slice(0, 10).join(', ')}\n\nFirst 5 rows:\n${JSON.stringify(jsonData.slice(0, 5), null, 2)}`;
    blendedPreview.textContent = blended.length > 0 ? JSON.stringify(blended.slice(0, 5), null, 2) : '⚠️ NO MATCHES - Lot numbers do not overlap between sources!';
    
    preview.style.display = 'block';
    
    status.textContent = `✓ Blended ${blended.length} rows! Review preview above, then click Download to proceed.`;
    
    // Add download button listener
    const downloadBtn = document.getElementById('downloadBtn');
    downloadBtn.style.display = 'block';
    downloadBtn.onclick = () => {
      const csv = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(blended));
      const dataUrl = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csv);
      chrome.downloads.download({
        url: dataUrl,
        filename: 'blended_data.csv'
      });
      status.textContent = '✓ Download started!';
      downloadBtn.style.display = 'none';
      preview.style.display = 'none';
    };
    
    return; // Stop here - don't download yet
  } catch (error) {
    console.error('Error:', error);
    status.textContent = 'Error: ' + error.message;
  }
});