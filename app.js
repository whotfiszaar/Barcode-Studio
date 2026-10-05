/**
 * Barcode Studio Pro - Universal Excel to Barcode PDF Generator
 * Fast, client-side, bug-free, mobile-responsive barcode utility.
 */

(function () {
  'use strict';

  // --- State Management ---
  const state = {
    records: [],
    columns: [],
    fileName: 'sample_locators.xlsx',
    activeSheet: '',
    currentPage: 1,
    totalPages: 1,
    zoom: 1.0,
    viewMode: 'sheet', // 'sheet' | 'card'
    exporting: false,
    cancelExport: false,
    
    // Layout parameters
    config: {
      preset: 'shahia_10',
      paperSize: 'a4', // 'a4' | 'letter' | 'legal' | 'custom'
      orientation: 'p', // 'p' | 'l'
      cols: 2,
      rows: 5,
      marginX: 8.8, // mm (approx 25pt)
      marginY: 12.3, // mm (approx 35pt)
      gapX: 8.8, // mm (approx 25pt)
      gapY: 5.3, // mm (approx 15pt)
      
      // Barcode & Card styling
      topField: 'ORG',
      topCustomText: '',
      barcodeField: 'LOCATIONS',
      bottomField: '__same__',
      barcodeType: 'code128',
      barHeight: 55, // pt
      headerSize: 14, // pt
      footerSize: 16, // pt
      borderStyle: 'rounded', // 'rounded' | 'square' | 'dashed' | 'none'
    }
  };

  // Preset Configurations
  const PRESETS = {
    shahia_10: {
      name: 'Shahia / Warehouse Standard (2x5)',
      paperSize: 'a4',
      orientation: 'p',
      cols: 2,
      rows: 5,
      marginX: 8.8,
      marginY: 12.3,
      gapX: 8.8,
      gapY: 5.3,
      borderStyle: 'rounded',
      barHeight: 55,
      headerSize: 14,
      footerSize: 16
    },
    compact_21: {
      name: 'Compact Shelf Tags (3x7)',
      paperSize: 'a4',
      orientation: 'p',
      cols: 3,
      rows: 7,
      marginX: 8,
      marginY: 10,
      gapX: 6,
      gapY: 4,
      borderStyle: 'rounded',
      barHeight: 38,
      headerSize: 11,
      footerSize: 12
    },
    pallet_6: {
      name: 'Pallet & Bin Cards (2x3)',
      paperSize: 'a4',
      orientation: 'p',
      cols: 2,
      rows: 3,
      marginX: 10,
      marginY: 15,
      gapX: 10,
      gapY: 10,
      borderStyle: 'rounded',
      barHeight: 80,
      headerSize: 18,
      footerSize: 22
    },
    avery_5160: {
      name: 'Avery 5160 Labels (3x10)',
      paperSize: 'letter',
      orientation: 'p',
      cols: 3,
      rows: 10,
      marginX: 4.8,
      marginY: 12.7,
      gapX: 3.2,
      gapY: 0,
      borderStyle: 'none',
      barHeight: 30,
      headerSize: 9,
      footerSize: 10
    },
    avery_5163: {
      name: 'Avery 5163 Shipping (2x5)',
      paperSize: 'letter',
      orientation: 'p',
      cols: 2,
      rows: 5,
      marginX: 4.0,
      marginY: 12.7,
      gapX: 3.5,
      gapY: 0,
      borderStyle: 'none',
      barHeight: 55,
      headerSize: 13,
      footerSize: 15
    },
    thermal_4x6: {
      name: 'Single 4"x6" Thermal Roll',
      paperSize: 'custom_4x6',
      orientation: 'p',
      cols: 1,
      rows: 1,
      marginX: 6,
      marginY: 6,
      gapX: 0,
      gapY: 0,
      borderStyle: 'rounded',
      barHeight: 90,
      headerSize: 22,
      footerSize: 26
    },
    thermal_3x2: {
      name: 'Single 3"x2" Shelf Sticker',
      paperSize: 'custom_3x2',
      orientation: 'l',
      cols: 1,
      rows: 1,
      marginX: 4,
      marginY: 4,
      gapX: 0,
      gapY: 0,
      borderStyle: 'rounded',
      barHeight: 50,
      headerSize: 14,
      footerSize: 16
    }
  };

  // Sample data to preload immediately on startup
  const SAMPLE_DATA = [
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-A-A' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-A-B' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-A-C' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-B-A' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-B-B' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-B-C' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-C-A' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-C-B' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-C-C' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-D-A' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-D-B' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-D-C' },
    { ORG: 'RUH_DC_2/RUH_MFG_MW_1', LOCATIONS: 'CR1-SL-E-A' },
    { ORG: 'RUH_DC_1', LOCATIONS: 'S1-SL-1A-A' },
    { ORG: 'RUH_DC_1', LOCATIONS: 'S1-SL-1A-B' },
    { ORG: 'DMM_DC_1', LOCATIONS: 'HA-FL-0A-A' },
    { ORG: 'ABH_DC_1', LOCATIONS: 'CR1-FL-VRL-1' },
    { ORG: 'SAK_DC_1', LOCATIONS: 'CR2-FL-VRL-1' },
    { ORG: 'TBK_DC_1', LOCATIONS: 'S1-FL-VRL-1' }
  ];

  // Helper conversions
  const MM_TO_PT = 72 / 25.4; // 1 mm = 2.83465 pt

  function getPageDimensionsPt() {
    const { paperSize, orientation } = state.config;
    let widthPt = 595.28; // Default A4
    let heightPt = 841.89;

    if (paperSize === 'letter') {
      widthPt = 612.0;
      heightPt = 792.0;
    } else if (paperSize === 'legal') {
      widthPt = 612.0;
      heightPt = 1008.0;
    } else if (paperSize === 'custom_4x6') {
      widthPt = 4 * 72; // 288 pt
      heightPt = 6 * 72; // 432 pt
    } else if (paperSize === 'custom_3x2') {
      widthPt = 3 * 72; // 216 pt
      heightPt = 2 * 72; // 144 pt
    }

    if (orientation === 'l') {
      return { width: Math.max(widthPt, heightPt), height: Math.min(widthPt, heightPt) };
    }
    return { width: Math.min(widthPt, heightPt), height: Math.max(widthPt, heightPt) };
  }

  // DOM Elements Cache
  const els = {};

  function initElements() {
    els.dropzone = document.getElementById('dropzone');
    els.fileInput = document.getElementById('fileInput');
    els.fileMeta = document.getElementById('fileMeta');
    els.fileNameLabel = document.getElementById('fileNameLabel');
    els.rowCountBadge = document.getElementById('rowCountBadge');
    els.sheetSelectGroup = document.getElementById('sheetSelectGroup');
    els.sheetSelect = document.getElementById('sheetSelect');
    
    els.topFieldSelect = document.getElementById('topFieldSelect');
    els.topCustomInput = document.getElementById('topCustomInput');
    els.barcodeFieldSelect = document.getElementById('barcodeFieldSelect');
    els.bottomFieldSelect = document.getElementById('bottomFieldSelect');
    
    els.templatePresetSelect = document.getElementById('templatePresetSelect');
    els.customLayoutControls = document.getElementById('customLayoutControls');
    els.paperSizeSelect = document.getElementById('paperSizeSelect');
    els.orientationSelect = document.getElementById('orientationSelect');
    els.gridColsInput = document.getElementById('gridColsInput');
    els.gridRowsInput = document.getElementById('gridRowsInput');
    els.marginXInput = document.getElementById('marginXInput');
    els.marginYInput = document.getElementById('marginYInput');
    els.gapXInput = document.getElementById('gapXInput');
    els.gapYInput = document.getElementById('gapYInput');
    
    els.barcodeTypeSelect = document.getElementById('barcodeTypeSelect');
    els.borderStyleSelect = document.getElementById('borderStyleSelect');
    els.barHeightSlider = document.getElementById('barHeightSlider');
    els.barHeightVal = document.getElementById('barHeightVal');
    els.headerSizeSlider = document.getElementById('headerSizeSlider');
    els.headerSizeVal = document.getElementById('headerSizeVal');
    els.footerSizeSlider = document.getElementById('footerSizeSlider');
    els.footerSizeVal = document.getElementById('footerSizeVal');
    
    els.paperSheet = document.getElementById('paperSheet');
    els.singleCardViewport = document.getElementById('singleCardViewport');
    els.singleCardWrapper = document.getElementById('singleCardWrapper');
    els.viewportWrapper = document.getElementById('viewportWrapper');
    
    els.btnViewSheet = document.getElementById('btnViewSheet');
    els.btnViewCard = document.getElementById('btnViewCard');
    els.btnPrevPage = document.getElementById('btnPrevPage');
    els.btnNextPage = document.getElementById('btnNextPage');
    els.pageNumberInput = document.getElementById('pageNumberInput');
    els.totalPagesLabel = document.getElementById('totalPagesLabel');
    
    els.zoomVal = document.getElementById('zoomVal');
    els.btnZoomIn = document.getElementById('btnZoomIn');
    els.btnZoomOut = document.getElementById('btnZoomOut');
    els.btnZoomFit = document.getElementById('btnZoomFit');
    
    els.statsSummary = document.getElementById('statsSummary');
    els.btnLoadSample = document.getElementById('btnLoadSample');
    els.btnPrintDirect = document.getElementById('btnPrintDirect');
    els.btnGeneratePdf = document.getElementById('btnGeneratePdf');
    els.btnQuickExport = document.getElementById('btnQuickExport');
    els.btnDownloadCurrentPage = document.getElementById('btnDownloadCurrentPage');
    
    els.progressModal = document.getElementById('progressModal');
    els.progressBarFill = document.getElementById('progressBarFill');
    els.progressStatusText = document.getElementById('progressStatusText');
    els.progressCountText = document.getElementById('progressCountText');
    els.progressPercentText = document.getElementById('progressPercentText');
    els.btnCancelExport = document.getElementById('btnCancelExport');
    
    els.mobileTabs = document.getElementById('mobileTabs');
    els.configSidebar = document.getElementById('configSidebar');
    els.previewArea = document.getElementById('previewArea');
  }

  // --- Setup Event Listeners ---
  function setupEvents() {
    // Dropzone & File Input
    els.dropzone.addEventListener('click', () => els.fileInput.click());
    els.fileInput.addEventListener('change', handleFileSelect);

    els.dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      els.dropzone.classList.add('dragover');
    });
    els.dropzone.addEventListener('dragleave', () => {
      els.dropzone.classList.remove('dragover');
    });
    els.dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      els.dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        processUploadedFile(e.dataTransfer.files[0]);
      }
    });

    // Sample data loader button
    els.btnLoadSample.addEventListener('click', loadSampleData);

    // Sheet switch
    els.sheetSelect.addEventListener('change', () => {
      if (window._currentWorkbook) {
        loadSheetData(window._currentWorkbook, els.sheetSelect.value);
      }
    });

    // Field mapping changes
    els.topFieldSelect.addEventListener('change', () => {
      state.config.topField = els.topFieldSelect.value;
      els.topCustomInput.style.display = els.topFieldSelect.value === '__custom__' ? 'block' : 'none';
      renderLivePreview();
    });
    els.topCustomInput.addEventListener('input', () => {
      state.config.topCustomText = els.topCustomInput.value;
      renderLivePreview();
    });
    els.barcodeFieldSelect.addEventListener('change', () => {
      state.config.barcodeField = els.barcodeFieldSelect.value;
      renderLivePreview();
    });
    els.bottomFieldSelect.addEventListener('change', () => {
      state.config.bottomField = els.bottomFieldSelect.value;
      renderLivePreview();
    });

    // Presets
    els.templatePresetSelect.addEventListener('change', () => {
      applyPreset(els.templatePresetSelect.value);
    });

    // Custom layout inputs
    const layoutInputs = [
      els.paperSizeSelect, els.orientationSelect,
      els.gridColsInput, els.gridRowsInput,
      els.marginXInput, els.marginYInput,
      els.gapXInput, els.gapYInput
    ];
    layoutInputs.forEach(input => {
      input.addEventListener('input', () => {
        readCustomLayoutFromUI();
        renderLivePreview();
      });
    });

    // Barcode type & Card style
    els.barcodeTypeSelect.addEventListener('change', () => {
      state.config.barcodeType = els.barcodeTypeSelect.value;
      renderLivePreview();
    });
    els.borderStyleSelect.addEventListener('change', () => {
      state.config.borderStyle = els.borderStyleSelect.value;
      renderLivePreview();
    });

    // Sliders
    els.barHeightSlider.addEventListener('input', () => {
      state.config.barHeight = parseInt(els.barHeightSlider.value, 10);
      els.barHeightVal.textContent = state.config.barHeight + 'pt';
      renderLivePreview();
    });
    els.headerSizeSlider.addEventListener('input', () => {
      state.config.headerSize = parseInt(els.headerSizeSlider.value, 10);
      els.headerSizeVal.textContent = state.config.headerSize + 'pt';
      renderLivePreview();
    });
    els.footerSizeSlider.addEventListener('input', () => {
      state.config.footerSize = parseInt(els.footerSizeSlider.value, 10);
      els.footerSizeVal.textContent = state.config.footerSize + 'pt';
      renderLivePreview();
    });

    // View mode toggles
    els.btnViewSheet.addEventListener('click', () => setViewMode('sheet'));
    els.btnViewCard.addEventListener('click', () => setViewMode('card'));

    // Pagination
    els.btnPrevPage.addEventListener('click', () => changePage(state.currentPage - 1));
    els.btnNextPage.addEventListener('click', () => changePage(state.currentPage + 1));
    els.pageNumberInput.addEventListener('change', () => {
      let p = parseInt(els.pageNumberInput.value, 10);
      if (isNaN(p)) p = 1;
      changePage(p);
    });

    // Zoom
    els.btnZoomIn.addEventListener('click', () => setZoom(state.zoom + 0.1));
    els.btnZoomOut.addEventListener('click', () => setZoom(state.zoom - 0.1));
    els.btnZoomFit.addEventListener('click', fitZoomToScreen);

    // Export & Print actions
    els.btnGeneratePdf.addEventListener('click', () => startPdfExport(false));
    els.btnQuickExport.addEventListener('click', () => startPdfExport(false));
    els.btnDownloadCurrentPage.addEventListener('click', () => startPdfExport(true));
    els.btnPrintDirect.addEventListener('click', () => window.print());
    els.btnCancelExport.addEventListener('click', () => { state.cancelExport = true; });

    // Mobile tabs switcher
    if (els.mobileTabs) {
      els.mobileTabs.addEventListener('click', (e) => {
        const btn = e.target.closest('.mobile-tab-btn');
        if (!btn) return;
        document.querySelectorAll('.mobile-tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const tab = btn.dataset.tab;
        if (tab === 'config') {
          els.configSidebar.classList.remove('tab-hidden');
          els.previewArea.classList.add('tab-hidden');
        } else {
          els.configSidebar.classList.add('tab-hidden');
          els.previewArea.classList.remove('tab-hidden');
          renderLivePreview();
        }
      });
    }

    // Window resize handler for zoom adaptation
    window.addEventListener('resize', debounce(() => {
      if (state.zoom === 1.0) fitZoomToScreen();
    }, 200));
  }

  // --- Preset Application ---
  function applyPreset(presetKey) {
    state.config.preset = presetKey;
    if (presetKey === 'custom') {
      els.customLayoutControls.style.display = 'flex';
      return;
    }

    const p = PRESETS[presetKey];
    if (!p) return;

    state.config.paperSize = p.paperSize;
    state.config.orientation = p.orientation;
    state.config.cols = p.cols;
    state.config.rows = p.rows;
    state.config.marginX = p.marginX;
    state.config.marginY = p.marginY;
    state.config.gapX = p.gapX;
    state.config.gapY = p.gapY;
    state.config.borderStyle = p.borderStyle;
    state.config.barHeight = p.barHeight;
    state.config.headerSize = p.headerSize;
    state.config.footerSize = p.footerSize;

    // Sync UI controls
    els.paperSizeSelect.value = p.paperSize.startsWith('custom') ? 'a4' : p.paperSize;
    els.orientationSelect.value = p.orientation;
    els.gridColsInput.value = p.cols;
    els.gridRowsInput.value = p.rows;
    els.marginXInput.value = p.marginX;
    els.marginYInput.value = p.marginY;
    els.gapXInput.value = p.gapX;
    els.gapYInput.value = p.gapY;
    els.borderStyleSelect.value = p.borderStyle;
    els.barHeightSlider.value = p.barHeight;
    els.barHeightVal.textContent = p.barHeight + 'pt';
    els.headerSizeSlider.value = p.headerSize;
    els.headerSizeVal.textContent = p.headerSize + 'pt';
    els.footerSizeSlider.value = p.footerSize;
    els.footerSizeVal.textContent = p.footerSize + 'pt';

    els.customLayoutControls.style.display = 'none';
    renderLivePreview();
  }

  function readCustomLayoutFromUI() {
    state.config.preset = 'custom';
    els.templatePresetSelect.value = 'custom';
    state.config.paperSize = els.paperSizeSelect.value;
    state.config.orientation = els.orientationSelect.value;
    state.config.cols = Math.max(1, parseInt(els.gridColsInput.value, 10) || 1);
    state.config.rows = Math.max(1, parseInt(els.gridRowsInput.value, 10) || 1);
    state.config.marginX = parseFloat(els.marginXInput.value) || 0;
    state.config.marginY = parseFloat(els.marginYInput.value) || 0;
    state.config.gapX = parseFloat(els.gapXInput.value) || 0;
    state.config.gapY = parseFloat(els.gapYInput.value) || 0;
  }

  // --- Excel File Loading ---
  function handleFileSelect(e) {
    const file = e.target.files && e.target.files[0];
    if (file) {
      processUploadedFile(file);
    }
  }

  function processUploadedFile(file) {
    state.fileName = file.name;
    els.fileNameLabel.textContent = file.name;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        window._currentWorkbook = workbook;

        // Populate sheet options
        els.sheetSelect.innerHTML = '';
        workbook.SheetNames.forEach(sheetName => {
          const opt = document.createElement('option');
          opt.value = sheetName;
          opt.textContent = sheetName;
          els.sheetSelect.appendChild(opt);
        });

        if (workbook.SheetNames.length > 1) {
          els.sheetSelectGroup.style.display = 'block';
        } else {
          els.sheetSelectGroup.style.display = 'none';
        }

        loadSheetData(workbook, workbook.SheetNames[0]);
      } catch (err) {
        alert('Error parsing Excel file. Please ensure it is a valid .xlsx, .xls, or .csv document.');
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function loadSheetData(workbook, sheetName) {
    state.activeSheet = sheetName;
    const worksheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (!rawRows || rawRows.length < 2) {
      alert('The selected sheet contains no data rows.');
      return;
    }

    // Extract headers (Row 0)
    const headerRow = rawRows[0];
    state.columns = headerRow.map((col, idx) => {
      const colStr = String(col || '').trim();
      return colStr || `Column_${idx + 1}`;
    });

    // Extract data rows
    state.records = [];
    for (let r = 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row || row.length === 0) continue;
      
      const record = {};
      let hasData = false;
      state.columns.forEach((colName, cIdx) => {
        const val = row[cIdx] !== undefined && row[cIdx] !== null ? String(row[cIdx]).trim() : '';
        record[colName] = val;
        if (val) hasData = true;
      });

      if (hasData) {
        state.records.push(record);
      }
    }

    els.rowCountBadge.textContent = `${state.records.length.toLocaleString()} records`;
    els.fileMeta.style.display = 'block';

    updateFieldMappingDropdowns();
    state.currentPage = 1;
    renderLivePreview();
  }

  function loadSampleData() {
    state.fileName = 'sample_warehouse_locators.xlsx';
    els.fileNameLabel.textContent = state.fileName;
    state.records = [...SAMPLE_DATA];
    state.columns = ['ORG', 'LOCATIONS'];
    els.rowCountBadge.textContent = `${state.records.length} records`;
    els.fileMeta.style.display = 'block';
    els.sheetSelectGroup.style.display = 'none';

    updateFieldMappingDropdowns();
    state.currentPage = 1;
    applyPreset('shahia_10');
  }

  function updateFieldMappingDropdowns() {
    // Populate Top Header Select
    els.topFieldSelect.innerHTML = '';
    const noneOpt = document.createElement('option');
    noneOpt.value = '__none__';
    noneOpt.textContent = 'None (Blank)';
    els.topFieldSelect.appendChild(noneOpt);

    const customOpt = document.createElement('option');
    customOpt.value = '__custom__';
    customOpt.textContent = 'Custom Fixed Text';
    els.topFieldSelect.appendChild(customOpt);

    state.columns.forEach(col => {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = col;
      els.topFieldSelect.appendChild(opt);
    });

    // Populate Barcode Value Select
    els.barcodeFieldSelect.innerHTML = '';
    state.columns.forEach(col => {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = col;
      els.barcodeFieldSelect.appendChild(opt);
    });

    // Populate Bottom Label Select
    els.bottomFieldSelect.innerHTML = '';
    const sameOpt = document.createElement('option');
    sameOpt.value = '__same__';
    sameOpt.textContent = 'Same as Barcode Value';
    els.bottomFieldSelect.appendChild(sameOpt);

    const hideOpt = document.createElement('option');
    hideOpt.value = '__none__';
    hideOpt.textContent = 'None (Hide)';
    els.bottomFieldSelect.appendChild(hideOpt);

    state.columns.forEach(col => {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = col;
      els.bottomFieldSelect.appendChild(opt);
    });

    // Smart auto-detection of column mapping
    const colLower = state.columns.map(c => c.toLowerCase());
    
    // Top Field: detect ORG / Facility / Organization / Company
    const orgIdx = colLower.findIndex(c => c.includes('org') || c.includes('facil') || c.includes('company'));
    if (orgIdx !== -1) {
      state.config.topField = state.columns[orgIdx];
      els.topFieldSelect.value = state.columns[orgIdx];
    } else {
      state.config.topField = state.columns[0];
      els.topFieldSelect.value = state.columns[0];
    }

    // Barcode Field: detect Barcode / Location / Loc / SKU / Code / Item
    const bcIdx = colLower.findIndex(c => c.includes('bar') || c.includes('loc') || c.includes('sku') || c.includes('item'));
    if (bcIdx !== -1 && bcIdx !== orgIdx) {
      state.config.barcodeField = state.columns[bcIdx];
      els.barcodeFieldSelect.value = state.columns[bcIdx];
    } else if (state.columns.length > 1) {
      state.config.barcodeField = state.columns[orgIdx === 0 ? 1 : 0];
      els.barcodeFieldSelect.value = state.config.barcodeField;
    } else {
      state.config.barcodeField = state.columns[0];
      els.barcodeFieldSelect.value = state.columns[0];
    }

    // Bottom Field: default to same as barcode
    state.config.bottomField = '__same__';
    els.bottomFieldSelect.value = '__same__';
  }

  // --- Rendering Engine (Live WYSIWYG) ---
  function renderLivePreview() {
    const { cols, rows, marginX, marginY, gapX, gapY, borderStyle } = state.config;
    const labelsPerPage = cols * rows;
    state.totalPages = Math.max(1, Math.ceil(state.records.length / labelsPerPage));

    if (state.currentPage > state.totalPages) {
      state.currentPage = state.totalPages;
    }

    // Update pagination labels
    els.pageNumberInput.value = state.currentPage;
    els.totalPagesLabel.textContent = state.totalPages;
    els.btnPrevPage.disabled = state.currentPage <= 1;
    els.btnNextPage.disabled = state.currentPage >= state.totalPages;

    // Update bottom status bar
    els.statsSummary.innerHTML = `
      <span class="dot"></span>
      Loaded <strong>${state.records.length.toLocaleString()}</strong> labels across <strong>${state.totalPages}</strong> pages (${labelsPerPage} per page).
    `;

    // Apply sheet paper dimensions
    const pageDim = getPageDimensionsPt();
    els.paperSheet.style.width = `${pageDim.width}px`;
    els.paperSheet.style.height = `${pageDim.height}px`;

    // Padding & Grid layout
    const padTop = marginY * MM_TO_PT;
    const padSide = marginX * MM_TO_PT;
    const gX = gapX * MM_TO_PT;
    const gY = gapY * MM_TO_PT;

    els.paperSheet.style.padding = `${padTop}px ${padSide}px`;
    els.paperSheet.style.columnGap = `${gX}px`;
    els.paperSheet.style.rowGap = `${gY}px`;
    els.paperSheet.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
    els.paperSheet.style.gridTemplateRows = `repeat(${rows}, 1fr)`;

    // Calculate items on current page
    const startIdx = (state.currentPage - 1) * labelsPerPage;
    const endIdx = Math.min(startIdx + labelsPerPage, state.records.length);
    const pageItems = state.records.slice(startIdx, endIdx);

    // Clear previous sheet cards
    els.paperSheet.innerHTML = '';

    // Render individual cards
    pageItems.forEach((record, localIdx) => {
      const card = createLabelCardElement(record, borderStyle);
      els.paperSheet.appendChild(card);
    });

    // Render Single Card Mode if active
    if (state.viewMode === 'card') {
      els.singleCardWrapper.innerHTML = '';
      if (pageItems.length > 0) {
        const sampleCard = createLabelCardElement(pageItems[0], borderStyle);
        els.singleCardWrapper.appendChild(sampleCard);
      }
    }
  }

  function createLabelCardElement(record, borderStyle) {
    const card = document.createElement('div');
    card.className = `label-card card-${borderStyle}`;

    // 1. Top Header Text
    let topText = '';
    if (state.config.topField === '__custom__') {
      topText = state.config.topCustomText;
    } else if (state.config.topField !== '__none__') {
      topText = record[state.config.topField] || '';
    }

    if (topText) {
      const topEl = document.createElement('div');
      topEl.className = 'label-top-text';
      topEl.style.fontSize = `${state.config.headerSize}px`;
      topEl.textContent = topText;
      card.appendChild(topEl);
    }

    // 2. Barcode Middle
    const barcodeVal = record[state.config.barcodeField] || '123456';
    const bcContainer = document.createElement('div');
    bcContainer.className = 'label-barcode-container';
    const canvas = document.createElement('canvas');
    bcContainer.appendChild(canvas);
    card.appendChild(bcContainer);

    renderBarcodeOnCanvas(canvas, barcodeVal, state.config.barcodeType, state.config.barHeight);

    // 3. Bottom Human Label
    let bottomText = '';
    if (state.config.bottomField === '__same__') {
      bottomText = barcodeVal;
    } else if (state.config.bottomField !== '__none__') {
      bottomText = record[state.config.bottomField] || '';
    }

    if (bottomText) {
      const bottomEl = document.createElement('div');
      bottomEl.className = 'label-bottom-text';
      bottomEl.style.fontSize = `${state.config.footerSize}px`;
      bottomEl.textContent = bottomText;
      card.appendChild(bottomEl);
    }

    return card;
  }

  function renderBarcodeOnCanvas(canvas, text, barcodeType, barHeightPt) {
    if (!window.bwipjs) {
      console.warn('bwipjs library not yet loaded');
      return;
    }

    try {
      const is2D = ['qrcode', 'datamatrix', 'pdf417', 'azteccode'].includes(barcodeType);
      
      const options = {
        bcid: barcodeType,
        text: String(text || '123456'),
        scale: 3,
        includetext: false
      };

      if (!is2D) {
        // Height in millimeters for 1D barcodes
        options.height = Math.max(10, Math.round(barHeightPt * 0.3528));
      }

      window.bwipjs.toCanvas(canvas, options);
    } catch (e) {
      // Graceful fallback for symbology constraints (e.g. non-numeric in EAN)
      const ctx = canvas.getContext('2d');
      canvas.width = 160;
      canvas.height = 40;
      ctx.fillStyle = '#fee2e2';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.fillStyle = '#991b1b';
      ctx.font = '10px sans-serif';
      ctx.fillText('Format Mismatch', 25, 24);
    }
  }

  // --- View Mode & Zooming ---
  function setViewMode(mode) {
    state.viewMode = mode;
    if (mode === 'sheet') {
      els.btnViewSheet.classList.add('active');
      els.btnViewCard.classList.remove('active');
      els.paperSheet.style.display = 'grid';
      els.singleCardViewport.style.display = 'none';
    } else {
      els.btnViewCard.classList.add('active');
      els.btnViewSheet.classList.remove('active');
      els.paperSheet.style.display = 'none';
      els.singleCardViewport.style.display = 'flex';
    }
    renderLivePreview();
  }

  function changePage(newPage) {
    if (newPage < 1 || newPage > state.totalPages) return;
    state.currentPage = newPage;
    renderLivePreview();
  }

  function setZoom(newZoom) {
    state.zoom = Math.max(0.3, Math.min(2.0, parseFloat(newZoom.toFixed(2))));
    els.zoomVal.textContent = Math.round(state.zoom * 100) + '%';
    els.paperSheet.style.transform = `scale(${state.zoom})`;
  }

  function fitZoomToScreen() {
    const wrapper = els.viewportWrapper;
    const availWidth = wrapper.clientWidth - 40;
    const availHeight = wrapper.clientHeight - 60;
    const pageDim = getPageDimensionsPt();

    const scaleX = availWidth / pageDim.width;
    const scaleY = availHeight / pageDim.height;
    const fitScale = Math.min(scaleX, scaleY, 1.0);
    setZoom(fitScale);
  }

  // --- High Quality jsPDF Vector/Canvas Export Engine ---
  async function startPdfExport(currentPageOnly = false) {
    if (state.records.length === 0) {
      alert('Please upload an Excel file or load sample data first.');
      return;
    }

    if (!window.jspdf || !window.jspdf.jsPDF) {
      alert('PDF generation engine is initializing. Please try again in a moment.');
      return;
    }

    const { jsPDF } = window.jspdf;
    const { cols, rows, marginX, marginY, gapX, gapY, borderStyle, barHeight, headerSize, footerSize } = state.config;
    const labelsPerPage = cols * rows;

    const totalPagesToExport = currentPageOnly ? 1 : Math.ceil(state.records.length / labelsPerPage);
    const startPage = currentPageOnly ? state.currentPage : 1;

    // Show Progress Modal
    state.exporting = true;
    state.cancelExport = false;
    els.progressModal.style.display = 'flex';
    els.progressBarFill.style.width = '0%';
    els.progressStatusText.textContent = 'Preparing vector document...';
    els.progressPercentText.textContent = '0%';
    els.progressCountText.textContent = `0 / ${totalPagesToExport} pages`;

    // Page format
    const pageDim = getPageDimensionsPt();
    const orientation = state.config.orientation === 'l' ? 'landscape' : 'portrait';

    // Initialize jsPDF document (unit = 'pt')
    const doc = new jsPDF({
      orientation: orientation,
      unit: 'pt',
      format: [pageDim.width, pageDim.height]
    });

    const padLeftPt = marginX * MM_TO_PT;
    const padTopPt = marginY * MM_TO_PT;
    const gapXPt = gapX * MM_TO_PT;
    const gapYPt = gapY * MM_TO_PT;

    const usableWidth = pageDim.width - 2 * padLeftPt - (cols - 1) * gapXPt;
    const usableHeight = pageDim.height - 2 * padTopPt - (rows - 1) * gapYPt;
    const boxWidth = usableWidth / cols;
    const boxHeight = usableHeight / rows;

    // Helper offscreen canvas for rendering barcodes
    const offscreenCanvas = document.createElement('canvas');

    for (let p = 0; p < totalPagesToExport; p++) {
      if (state.cancelExport) {
        break;
      }

      const activePageNum = startPage + p;
      if (p > 0) {
        doc.addPage([pageDim.width, pageDim.height], orientation);
      }

      const pageStartIdx = (activePageNum - 1) * labelsPerPage;
      const pageEndIdx = Math.min(pageStartIdx + labelsPerPage, state.records.length);
      const pageRecords = state.records.slice(pageStartIdx, pageEndIdx);

      // Render cards for this page
      for (let i = 0; i < pageRecords.length; i++) {
        const record = pageRecords[i];
        const colIdx = i % cols;
        const rowIdx = Math.floor(i / cols);

        const cardX = padLeftPt + colIdx * (boxWidth + gapXPt);
        const cardY = padTopPt + rowIdx * (boxHeight + gapYPt);
        const centerX = cardX + boxWidth / 2.0;

        // 1. Draw Card Border
        doc.setLineWidth(1.0);
        doc.setDrawColor(51, 51, 51); // 0.2 gray

        if (borderStyle === 'rounded') {
          doc.roundedRect(cardX, cardY, boxWidth, boxHeight, 4, 4, 'S');
        } else if (borderStyle === 'square') {
          doc.rect(cardX, cardY, boxWidth, boxHeight, 'S');
        } else if (borderStyle === 'dashed') {
          doc.setLineDashPattern([3, 3], 0);
          doc.rect(cardX, cardY, boxWidth, boxHeight, 'S');
          doc.setLineDashPattern([], 0);
        }

        // 2. Top Header Text
        let topText = '';
        if (state.config.topField === '__custom__') {
          topText = state.config.topCustomText;
        } else if (state.config.topField !== '__none__') {
          topText = record[state.config.topField] || '';
        }

        if (topText) {
          doc.setFont('helvetica', 'bold');
          doc.setFontSize(headerSize);
          doc.setTextColor(0, 0, 0);
          // Y baseline offset from top of card
          doc.text(String(topText), centerX, cardY + 20, { align: 'center' });
        }

        // 3. Middle Barcode
        const barcodeVal = record[state.config.barcodeField] || '123456';
        renderBarcodeOnCanvas(offscreenCanvas, barcodeVal, state.config.barcodeType, barHeight);

        try {
          const barcodeImgData = offscreenCanvas.toDataURL('image/png');
          
          // Fit barcode nicely inside card
          const maxBcW = boxWidth - 20.0;
          const maxBcH = boxHeight - (topText ? 30 : 15) - 30;
          
          let renderW = (offscreenCanvas.width / 3.0);
          let renderH = (offscreenCanvas.height / 3.0);

          if (renderW > maxBcW) {
            const ratio = maxBcW / renderW;
            renderW = maxBcW;
            renderH = renderH * ratio;
          }
          if (renderH > maxBcH) {
            const ratio = maxBcH / renderH;
            renderH = maxBcH;
            renderW = renderW * ratio;
          }

          const bcX = centerX - renderW / 2.0;
          const bcY = cardY + (boxHeight - renderH) / 2.0 + (topText ? 2 : -5);

          doc.addImage(barcodeImgData, 'PNG', bcX, bcY, renderW, renderH);
        } catch (imgErr) {
          console.error('Failed to add barcode image to PDF', imgErr);
        }

        // 4. Bottom Human Text
        let bottomText = '';
        if (state.config.bottomField === '__same__') {
          bottomText = barcodeVal;
        } else if (state.config.bottomField !== '__none__') {
          bottomText = record[state.config.bottomField] || '';
        }

        if (bottomText) {
          doc.setFont('helvetica', 'bold');
          doc.setFontSize(footerSize);
          doc.setTextColor(0, 0, 0);
          doc.text(String(bottomText), centerX, cardY + boxHeight - 12, { align: 'center' });
        }
      }

      // Update progress
      const percent = Math.round(((p + 1) / totalPagesToExport) * 100);
      els.progressBarFill.style.width = `${percent}%`;
      els.progressPercentText.textContent = `${percent}%`;
      els.progressCountText.textContent = `${p + 1} / ${totalPagesToExport} pages`;

      // Yield thread so UI and animation update smoothly
      await new Promise(r => setTimeout(r, 10));
    }

    if (!state.cancelExport) {
      els.progressStatusText.textContent = 'Saving PDF file...';
      const cleanName = (state.fileName || 'Barcodes').replace(/\.[^/.]+$/, '');
      const outName = currentPageOnly 
        ? `2D_Barcode_${cleanName}_page_${state.currentPage}.pdf`
        : `2D_Barcode_${cleanName}.pdf`;

      doc.save(outName);
    }

    state.exporting = false;
    els.progressModal.style.display = 'none';
  }

  // --- Utility: Debounce ---
  function debounce(fn, ms) {
    let timer;
    return function (...args) {
      clearTimeout(timer);
      timer = setTimeout(() => fn.apply(this, args), ms);
    };
  }

  // --- Initialize Application on Page Load ---
  document.addEventListener('DOMContentLoaded', () => {
    initElements();
    setupEvents();
    loadSampleData();
    fitZoomToScreen();
  });

})();
