const state = {
  currentFile: '',
  sheets: [],
  currentSheet: '',
  headers: [],
  rows: [],
  visibleHeaders: [],
  selectedRowNumber: null,
  staff: [],
  clsList: [],
  focusedInput: null,
  rowChanges: {},
  batchSaveTimer: null,
  saveQueue: Promise.resolve(),
  newRowValues: null,
  autosaveDraftTimer: null,
  lastDraftSignature: '',
  suggestItems: [],
  suggestActiveIndex: -1,
  suppressNextSuggestion: false,
  sheetMeta: {},
  clearPreviewData: null,
  dismissedSampleNoticeFor: '',
  emrConfig: null,
  pendingEmrAction: '',
  headerAliases: {},
  soPhauThuat: { danhSachFile: '', headers: [], rows: [], filtered: [], selectedIndex: -1, imagesLoaded: false },
  soPtDoiChieu: [],
  soPtDoiChieuOpenRows: new Set(),
  documentAiConfig: null,
  documentAiRawOcrFile: '',
};

const $ = (id) => document.getElementById(id);
const AUTOSAVE_KEY = 'pm_ctch_t5_inline_autosave_v1';
const SURGERY_ENTRY_HEADER_KEYS = new Set(['ptvchinh', 'phumo1', 'phumo2', 'phumo3']);

function resetDocumentHorizontalScroll() {
  // Bảng có thanh cuộn ngang riêng; trang chính luôn phải đứng ở mép trái.
  document.documentElement.scrollLeft = 0;
  if (document.body) document.body.scrollLeft = 0;
}



function showToast(message, tone = 'success', timeout = 3600) {
  const host = $('toastHost');
  if (!host) return;
  const item = document.createElement('div');
  item.className = `toast ${tone}`;
  item.textContent = message;
  host.appendChild(item);
  setTimeout(() => item.remove(), timeout);
}

function extractPeriodFromFile(file) {
  const name = String(file || '');
  const monthMatch = name.match(/(?:^|[^a-z0-9])T(?:HANG)?\s*0?([1-9]|1[0-2])(?:[^0-9]|$)/i)
    || name.match(/THA(?:N|NG)G?\s*0?([1-9]|1[0-2])/i);
  const yearMatch = name.match(/(20\d{2})/);
  return {
    month: monthMatch ? Number(monthMatch[1]) : null,
    year: yearMatch ? Number(yearMatch[1]) : null,
  };
}

function inferPeriodFromRows() {
  const dateHeader = findHeader(['NGÀY', 'Ngày']);
  if (!dateHeader) return { month: null, year: null };
  const counts = new Map();
  state.rows.slice(0, 200).forEach(row => {
    const text = String(row.values?.[dateHeader] ?? '').trim();
    const match = text.match(/(?:^|\s)([0-3]?\d)[\/-]([01]?\d)[\/-](20\d{2})(?:\s|$)/);
    if (!match) return;
    const month = Number(match[2]);
    const year = Number(match[3]);
    if (month < 1 || month > 12) return;
    const key = `${month}/${year}`;
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  const best = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
  if (!best) return { month: null, year: null };
  const [month, year] = best[0].split('/').map(Number);
  return { month, year };
}

function workbookLabel(file = state.currentFile) {
  const fromFile = extractPeriodFromFile(file);
  const fromRows = inferPeriodFromRows();
  const month = fromFile.month || fromRows.month;
  const year = fromFile.year || fromRows.year;
  if (month && year) return `Tháng ${month}/${year}`;
  if (month) return `Tháng ${month}`;
  if (year) return `Năm ${year}`;
  return 'File đang làm việc';
}

function updateWorkspaceIdentity() {
  const periodLabel = workbookLabel();
  const kind = sheetKind();
  const kindLabel = kind === 'surgery' ? 'Phẫu thuật' : (kind === 'procedure' ? 'Thủ thuật – tiểu phẫu' : 'Nhập liệu');
  const fileName = String(state.currentFile || '').split(/[\\/]/).pop() || '';
  const pageTitle = $('pageTitle');
  const pageSubtitle = $('pageSubtitle');
  if (pageTitle) pageTitle.textContent = `${kindLabel} · ${periodLabel}`;
  if (pageSubtitle) pageSubtitle.textContent = fileName
    ? `${fileName} · Sheet ${state.currentSheet || 'chưa chọn'}`
    : 'Chọn file tháng cần làm việc để bắt đầu.';
  document.title = `${kindLabel} · ${periodLabel} | PM CTCH`;
}

function assistantHeaderMap() {
  const candidates = {
    phuMo1: ['Phụ mổ 1'],
    phuMo2: ['Phụ mổ 2'],
    phuMo3: ['Phụ mổ 3'],
    ptvChinh: ['PTV chính'],
  };
  const result = {};
  Object.entries(candidates).forEach(([key, names]) => {
    result[key] = findHeader(names);
  });
  return result;
}

function countNonBlankInHeader(header) {
  if (!header) return 0;
  return state.rows.reduce((count, row) => {
    const value = state.rowChanges[row.rowNumber]?.[header] ?? row.values?.[header];
    return count + (String(value ?? '').trim() ? 1 : 0);
  }, 0);
}

function updateDirtyMetrics() {
  const dirtyCount = Object.keys(state.rowChanges || {}).length + (state.newRowValues ? 1 : 0);
  const chip = $('dirtyStatus');
  if (chip) {
    chip.textContent = dirtyCount ? `${dirtyCount} dòng chưa hoàn tất` : '0 thay đổi';
    chip.classList.toggle('dirty', dirtyCount > 0);
    chip.classList.toggle('neutral', dirtyCount === 0);
  }
}

function updateDashboardMetrics() {
  const rowCount = state.sheetMeta.matchedRowCount ?? state.sheetMeta.rowCount ?? state.rows.length;
  const missingCount = state.sheetMeta.missingCount ?? state.rows.filter(row => (row.missing || []).length > 0).length;
  if ($('rowCountMetric')) $('rowCountMetric').textContent = String(rowCount || 0);
  if ($('missingCountMetric')) $('missingCountMetric').textContent = String(missingCount || 0);
  updateDirtyMetrics();
  updateDataQualityPanel();
}

function updateDataQualityPanel() {
  const summary = $('qualitySummary');
  const panel = $('qualityPanel');
  if (!summary || !panel) return;

  const total = Number(state.sheetMeta.matchedRowCount ?? state.rows.length ?? 0);
  const complete = Number(state.sheetMeta.completeCount ?? state.rows.filter(row => !(row.missing || []).length).length);
  const missing = Math.max(0, total - complete);
  const missingByHeader = state.sheetMeta.missingByHeader || {};
  const blankRows = Number(state.sheetMeta.blankRows || 0);
  const templateRows = Number(state.sheetMeta.templateRows || 0);
  const slots = Number(state.sheetMeta.dataSlots || 0);

  if (!state.currentSheet || !total) {
    summary.textContent = 'Kiểm tra dữ liệu';
    panel.innerHTML = '<span>Chưa có dữ liệu để đánh giá.</span>';
    return;
  }

  summary.textContent = missing ? `${complete}/${total} dòng đủ` : `${total}/${total} dòng đủ`;
  const missingItems = Object.entries(missingByHeader)
    .filter(([, count]) => Number(count) > 0)
    .sort((a, b) => Number(b[1]) - Number(a[1]))
    .map(([header, count]) => `<li><strong>${escapeHtml(header)}</strong>: thiếu ${Number(count)} dòng</li>`)
    .join('');

  const notes = [];
  if (sheetKind() === 'surgery' && Number(missingByHeader['PTV chính'] || 0) === total && total > 0) {
    notes.push('Toàn bộ dòng đang thiếu PTV chính. Sau khi xóa dữ liệu mẫu phụ mổ, đây là trạng thái bình thường trước khi nhập nhân sự thật.');
  }
  if (blankRows || templateRows) {
    notes.push(`${blankRows + templateRows} vị trí giữa tiêu đề và Tổng Cộng không được đưa vào bảng vì là dòng trống hoặc chỉ có công thức mẫu.`);
  }
  if (slots && total > slots) {
    notes.push('Số dòng nhận diện lớn hơn vùng dữ liệu dự kiến; nên kiểm tra lại dòng Tổng Cộng.');
  }

  panel.innerHTML = `
    <h4>Đánh giá sheet ${escapeHtml(state.currentSheet)}</h4>
    <div class="${missing ? 'quality-warn' : 'quality-good'}">${missing ? `${missing} dòng cần bổ sung` : 'Các cột bắt buộc đã đầy đủ'}</div>
    ${missingItems ? `<ul class="quality-list">${missingItems}</ul>` : ''}
    ${notes.length ? `<ul class="quality-list">${notes.map(note => `<li>${escapeHtml(note)}</li>`).join('')}</ul>` : ''}
  `;
}

function updateContextUI() {
  const kind = sheetKind();
  const subtype = sheetSubtype();
  document.querySelectorAll('[data-context="surgery"]').forEach(el => el.classList.toggle('hidden', subtype !== 'surgery'));
  document.querySelectorAll('[data-context="thuthuat"]').forEach(el => el.classList.toggle('hidden', subtype !== 'thuthuat'));
  document.querySelectorAll('[data-context="tieuphau"]').forEach(el => el.classList.toggle('hidden', subtype !== 'tieuphau'));

  // File phẫu thuật đã có sẵn danh sách bệnh nhân. Ẩn thao tác thêm/xóa dòng để
  // người nhập chỉ tập trung vào PTV chính và Phụ mổ 1–3.
  const rowActions = document.querySelector('.command-primary');
  const commandDivider = document.querySelector('.command-divider');
  if (rowActions) rowActions.classList.toggle('hidden', kind === 'surgery');
  if (commandDivider) commandDivider.classList.toggle('hidden', kind === 'surgery');
  const keyboardHint = $('keyboardHint');
  if (keyboardHint) {
    keyboardHint.textContent = kind === 'surgery'
      ? 'Tab: sang cột kế · Enter: xuống dòng · Shift: đi ngược · Mũi tên: di chuyển'
      : 'Enter: xuống dòng · Tab: sang ô kế';
  }

  const labels = {
    surgery: {
      title: 'Công cụ nhập liệu phẫu thuật',
      subtitle: 'Chỉ nhập PTV chính và Phụ mổ 1–3; các cột còn lại được khóa để bảo vệ dữ liệu.',
      hint: 'Chỉ bốn cột PTV chính và Phụ mổ 1–3 có thể chỉnh sửa.',
      badge: 'Phẫu thuật',
    },
    procedure: {
      title: 'Công cụ thủ thuật – tiểu phẫu',
      subtitle: 'Chuyển các kỹ thuật phù hợp sang tiểu phẫu và bổ sung bác sĩ.',
      hint: 'Sheet thủ thuật/tiểu phẫu: hiển thị kỹ thuật, số lượng và bác sĩ.',
      badge: 'Thủ thuật – tiểu phẫu',
    },
    tieuphau: {
      title: 'Nhập liệu tiểu phẫu',
      subtitle: 'Đối chiếu thời gian trên EMR và tự điền bí danh Bác sĩ thực hiện.',
      hint: 'Cột Bác sĩ được đối chiếu từ D/s Thủ thuật và D/s Phẫu thuật.',
      badge: 'Tiểu phẫu · EMR',
    },
    compact: {
      title: 'Công cụ nhập liệu',
      subtitle: 'Chọn chế độ hiển thị hoặc sheet chuẩn để dùng công cụ tự động.',
      hint: 'Sheet chưa nhận diện; chương trình đang dùng chế độ hiển thị gọn.',
      badge: 'Sheet khác',
    },
  };
  const current = labels[subtype === 'tieuphau' ? 'tieuphau' : kind] || labels.compact;
  if ($('actionTitle')) $('actionTitle').textContent = current.title;
  if ($('actionSubtitle')) $('actionSubtitle').textContent = current.subtitle;
  if ($('sheetContextHint')) $('sheetContextHint').textContent = current.hint;
  if ($('tableTitle')) $('tableTitle').textContent = state.currentSheet ? `Dữ liệu · ${state.currentSheet}` : 'Dữ liệu';
  const badge = $('sheetTypeBadge');
  if (badge) {
    badge.textContent = current.badge;
    badge.className = `type-badge ${kind}`;
  }
  updateWorkspaceIdentity();
  updateSampleDataNotice();
}

function updateSampleDataNotice() {
  const notice = $('sampleDataNotice');
  if (!notice) return;
  const headers = assistantHeaderMap();
  const columns = [headers.phuMo1, headers.phuMo2, headers.phuMo3].filter(Boolean);
  const count = columns.reduce((sum, header) => sum + countNonBlankInHeader(header), 0);
  const noticeKey = `${state.currentFile}|${state.currentSheet}`;
  const isSourceFile = state.currentFile && !/_NHAP_LIEU\.(xlsx|xlsm)$/i.test(state.currentFile);
  const shouldShow = sheetKind() === 'surgery' && columns.length > 0 && count > 0 && isSourceFile && state.dismissedSampleNoticeFor !== noticeKey;
  notice.classList.toggle('hidden', !shouldShow);
  if ($('sampleDataNoticeText')) {
    $('sampleDataNoticeText').textContent = `${count} ô Phụ mổ 1–3 đang có giá trị trong file nguồn. Nếu đây là mẫu tháng trước, hãy xóa trước khi bổ sung dữ liệu thật.`;
  }
}

function log(message) {
  const now = new Date().toLocaleTimeString('vi-VN');
  const logBox = $('logBox');
  if (logBox) logBox.textContent = `[${now}] ${message}\n` + logBox.textContent;
}

function setBusy(isBusy, text = 'Đang xử lý...') {
  document.querySelectorAll('button').forEach(btn => {
    if (!btn.classList.contains('cell-btn') && !btn.classList.contains('staff-chip') && !btn.classList.contains('row-action-btn')) {
      btn.disabled = isBusy;
    }
  });
  if (isBusy) log(text);
}

async function api(url, options = {}) {
  const res = await fetch(url, options);
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.ok === false) {
    throw new Error(data.error || `Lỗi HTTP ${res.status}`);
  }
  return data;
}

function autosaveStatus(message, tone = 'normal') {
  const el = $('autosaveStatus');
  if (!el) return;
  el.textContent = message;
  el.classList.remove('saving', 'ok', 'error', 'draft');
  if (tone) el.classList.add(tone);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function normalizeKey(value) {
  return String(value ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/[^a-z0-9]+/g, '');
}

function displayPath(file) {
  return String(file || '').split(/[\\/]/).slice(-2).join('/');
}

function downloadLink(file, label = 'Tải file') {
  return `<a class="btn light" target="_blank" href="/download?finalize=1&file=${encodeURIComponent(file)}">${label}</a>`;
}

function showResult(el, data) {
  if (!el) return;
  el.classList.remove('hidden');
  const file = data.file || data.file_dau_ra;
  const lines = [];
  if (data.message) lines.push(data.message);
  if (file) lines.push(`File làm việc: ${displayPath(file)}`);

  const labelMap = {
    so_dong_chuyen: 'Dòng chuyển sang tiểu phẫu',
    so_dong_dien_bs: 'Dòng đã điền BS',
    so_dong_chua_dien_bs: 'Dòng chưa điền BS',
    rowNumber: 'Dòng Excel',
    deletedRow: 'Dòng đã xóa',
    checkedRows: 'Dòng cần kiểm tra EMR',
    updatedRows: 'Dòng đã điền Bác sĩ',
    skippedValidRows: 'Dòng đã đúng Bác sĩ',
    unmatchedRows: 'Dòng chưa tìm được'
  };
  Object.entries(data).forEach(([key, value]) => {
    if (Object.prototype.hasOwnProperty.call(labelMap, key) && typeof value !== 'object') {
      lines.push(`${labelMap[key]}: ${value}`);
    }
  });
  if (data.report || data.file_bao_cao) lines.push('Báo cáo phụ đã được ghi đè lại, không tạo nhiều file kết quả.');
  if (Array.isArray(data.unmatched) && data.unmatched.length) {
    lines.push('Các dòng cần kiểm tra thủ công:');
    data.unmatched.slice(0, 20).forEach(item => {
      const row = item.rowNumber ? `Dòng ${item.rowNumber}` : 'Dòng chưa xác định';
      const patient = item.patient ? ` · ${item.patient}` : '';
      const reason = item.reason ? ` — ${item.reason}` : '';
      lines.push(`${row}${patient}${reason}`);
    });
    if (data.unmatched.length > 20) lines.push(`... còn ${data.unmatched.length - 20} dòng khác.`);
  }

  const summaryText = lines[0] || 'Tác vụ đã hoàn tất';
  const detailLines = lines.slice(1);
  const hasDetails = detailLines.length > 0 || Boolean(file);
  el.innerHTML = `
    <details class="operation-details">
      <summary>
        <span class="operation-status-icon" aria-hidden="true">✓</span>
        <strong>${escapeHtml(summaryText)}</strong>
        ${hasDetails ? '<span class="operation-expand-label">Xem chi tiết</span>' : ''}
      </summary>
      ${hasDetails ? `<div class="operation-detail-body">
        ${detailLines.length ? `<pre>${escapeHtml(detailLines.join('\n'))}</pre>` : ''}
        <div class="result-actions">${file ? downloadLink(file, 'Tải file Excel') : ''}</div>
      </div>` : ''}
    </details>`;
}

function isStaffHeader(header) {
  const key = normalizeKey(header);
  return key === 'bs'
    || key === 'bacsi'
    || key.includes('phumo')
    || key.includes('ptv')
    || key.includes('phauthuatvien')
    || key.includes('dieuduong')
    || key === 'ktv'
    || key.includes('kythuatvien');
}

function isNameHeader(header) {
  const key = normalizeKey(header);
  return key.includes('hovaten');
}

function isLongTextHeader(header) {
  const key = normalizeKey(header);
  return key.includes('phuongphap') || key.includes('chandoan') || key.includes('tencls');
}

function isDateHeader(header) {
  return normalizeKey(header).includes('ngay');
}

function isNumberHeader(header) {
  const key = normalizeKey(header);
  return key === 'tuoi' || key.includes('soluong') || key.includes('thanhtien');
}

function isFormulaLike(value) {
  return typeof value === 'string' && value.trim().startsWith('=');
}

function isLikelyFormulaHeader(header) {
  const key = normalizeKey(header);
  return key.startsWith('sotien') || key === 'thuclanh' || key.includes('congthuc');
}

function isSurgeryEntryHeader(header) {
  return SURGERY_ENTRY_HEADER_KEYS.has(normalizeKey(header));
}

function isReadOnlyHeader(header) {
  const key = normalizeKey(header);
  // Ở sheet phẫu thuật, chương trình chỉ cho phép nhập bốn cột nhân sự.
  // Các cột bệnh nhân/kỹ thuật chỉ dùng để đối chiếu và tuyệt đối không ghi lại.
  if (sheetKind() === 'surgery') return !isSurgeryEntryHeader(header);
  return key === 'stt' || isLikelyFormulaHeader(header);
}

function sanitizeChangesForCurrentSheet(changes) {
  if (sheetKind() !== 'surgery') return changes || {};
  const safe = {};
  Object.entries(changes || {}).forEach(([rowNumber, rowData]) => {
    const filtered = {};
    Object.entries(rowData || {}).forEach(([header, value]) => {
      if (isSurgeryEntryHeader(header)) filtered[header] = value;
    });
    if (Object.keys(filtered).length) safe[rowNumber] = filtered;
  });
  return safe;
}

function headerExists(header) {
  return state.headers.some(h => normalizeKey(h) === normalizeKey(header));
}

function findHeader(candidates) {
  for (const candidate of candidates) {
    const hit = state.headers.find(h => normalizeKey(h) === normalizeKey(candidate));
    if (hit) return hit;
  }
  return null;
}

function sheetKind() {
  const key = normalizeKey(state.currentSheet);
  if (key.includes('phauthuat')) return 'surgery';
  if (key.includes('thuthuat') || key.includes('tieuphau')) return 'procedure';
  return 'compact';
}

function sheetSubtype() {
  const key = normalizeKey(state.currentSheet);
  if (key.includes('phauthuat')) return 'surgery';
  if (key.includes('tieuphau')) return 'tieuphau';
  if (key.includes('thuthuat')) return 'thuthuat';
  return 'compact';
}

function currentViewMode() {
  const selected = $('viewModeSelect') ? $('viewModeSelect').value : 'auto';
  return selected === 'auto' ? sheetKind() : selected;
}

function headerFormulaRate(header) {
  let seen = 0;
  let formulas = 0;
  state.rows.slice(0, 80).forEach(row => {
    const value = row.values[header];
    if (value !== null && value !== undefined && String(value).trim() !== '') {
      seen += 1;
      if (isFormulaLike(value)) formulas += 1;
    }
  });
  return seen ? formulas / seen : 0;
}

function shouldHideHeader(header) {
  const key = normalizeKey(header);
  if (key === 'loaiphauthuat') return true;
  if (isLikelyFormulaHeader(header)) return true;
  if (headerFormulaRate(header) > 0.5) return true;
  return false;
}

function aliasCandidates(canonicalName) {
  const configured = state.headerAliases && state.headerAliases[canonicalName];
  return configured && configured.length ? configured : [canonicalName];
}

function chooseDisplayHeaders() {
  const mode = currentViewMode();
  const surgery = [
    ['STT'],
    aliasCandidates('Ngày'),
    aliasCandidates('Họ và tên'),
    aliasCandidates('Tuổi'),
    aliasCandidates('Tên CLS'),
    ['PTV chính'],
    ['Phụ mổ 1'],
    ['Phụ mổ 2'],
    ['Phụ mổ 3']
  ];
  const procedure = [
    ['STT'],
    aliasCandidates('Ngày'),
    aliasCandidates('Họ và tên'),
    aliasCandidates('Tuổi'),
    aliasCandidates('Tên CLS'),
    aliasCandidates('Số Lượng'),
    ['Thành tiền'],
    aliasCandidates('Bác Sĩ'),
    aliasCandidates('Điều Dưỡng'),
    ['Ghi chú']
  ];
  const preferred = mode === 'surgery' ? surgery : (mode === 'procedure' ? procedure : [...surgery, ...procedure]);
  const picked = [];
  preferred.forEach(group => {
    const hit = findHeader(group);
    if (hit && !picked.includes(hit) && !shouldHideHeader(hit)) picked.push(hit);
  });

  if (mode === 'compact') {
    state.headers.forEach(h => {
      if (picked.length < 15 && !picked.includes(h) && !shouldHideHeader(h)) picked.push(h);
    });
  }

  if (!picked.length) {
    state.headers.forEach(h => {
      if (picked.length < 12 && !shouldHideHeader(h)) picked.push(h);
    });
  }
  return picked;
}

function cellClassForHeader(header) {
  const classes = [];
  if (isLongTextHeader(header)) classes.push('cell-text-long');
  if (isNameHeader(header)) classes.push('cell-name');
  if (isDateHeader(header)) classes.push('cell-date');
  if (isStaffHeader(header)) classes.push('cell-staff');
  if (isNumberHeader(header)) classes.push('cell-number');
  return classes.join(' ');
}

function dataCellClass(header, index) {
  const classes = cellClassForHeader(header).split(/\s+/).filter(Boolean);
  if (index === 0) classes.push('sticky-index');
  if (isNameHeader(header)) classes.push('sticky-name');
  if (sheetKind() === 'surgery') {
    classes.push(isSurgeryEntryHeader(header) ? 'entry-column' : 'context-column');
  }
  return [...new Set(classes)].join(' ');
}

function cellInputClass(header, changed, readonly) {
  const classes = ['cell-input'];
  if (isStaffHeader(header)) classes.push('staff-cell');
  if (sheetKind() === 'surgery' && isSurgeryEntryHeader(header)) classes.push('entry-cell');
  if (changed) classes.push('changed');
  if (readonly) classes.push('readonly');
  return classes.join(' ');
}

function inputDatalistId(_header) {
  return '';
}


function uniqueColumnValues(header, limit = 250) {
  const seen = new Set();
  const values = [];
  state.rows.forEach(row => {
    const value = row.values[header];
    const text = String(value ?? '').trim();
    if (!text || isFormulaLike(text)) return;
    const key = normalizeKey(text);
    if (seen.has(key)) return;
    seen.add(key);
    values.push(text);
  });
  return values.slice(0, limit);
}

function renderDatalists() {
  // Không dùng datalist native vì Chrome mở khung gợi ý rất lớn và gây lag.
}


function renderStaffDatalist() {
  // Gợi ý nhân sự dùng khung riêng nhỏ gọn, chỉ hiện khi người dùng gõ.
}


function renderStaffQuickList() {
  const box = $('staffQuickBox');
  if (!box) return;
  const bsnt = state.staff.filter(item => item.vaiTro === 'bsnt');
  const groups = {};
  bsnt.forEach(item => {
    const group = item.nhom || 'BSNT';
    if (!groups[group]) groups[group] = [];
    groups[group].push(item);
  });
  const groupNames = Object.keys(groups);
  box.innerHTML = groupNames.map(group => {
    const chips = groups[group].map(item =>
      `<button type="button" class="staff-chip" data-alias="${escapeHtml(item.biDanh)}" title="${escapeHtml(item.hoTen)}">${escapeHtml(item.biDanh)}</button>`
    ).join('');
    return `<div class="staff-group"><div class="staff-group-title">${escapeHtml(group)}</div><div class="staff-chip-row">${chips}</div></div>`;
  }).join('');
  box.querySelectorAll('[data-alias]').forEach(btn => {
    btn.addEventListener('click', () => fillFocusedStaff(btn.dataset.alias));
  });
}

function fillFocusedStaff(alias) {
  const input = state.focusedInput;
  if (!input || !document.body.contains(input) || input.disabled || input.readOnly) {
    alert('Bấm vào ô PTV chính / Phụ mổ / BS trước, rồi chọn bí danh BSNT.');
    return;
  }
  input.value = alias;
  input.focus();
  input.dispatchEvent(new Event('input', { bubbles: true }));
  hideSuggestPanel();
}
function staffMatchScore(item, queryKey) {
  const aliasKey = normalizeKey(item.biDanh);
  const nameKey = normalizeKey(item.hoTen);
  if (!queryKey) return 999;
  if (aliasKey === queryKey) return 0;
  if (aliasKey.startsWith(queryKey)) return 1;
  if (nameKey.startsWith(queryKey)) return 2;
  if (aliasKey.includes(queryKey)) return 3;
  if (nameKey.includes(queryKey)) return 4;
  return 999;
}

function allowedStaffRolesForInput(input) {
  const headerKey = normalizeKey(input?.dataset?.header || '');
  const kind = sheetKind();

  // PTV và phụ mổ trong sheet phẫu thuật chỉ là bác sĩ/BSNT.
  if (kind === 'surgery') return ['bac_si', 'bsnt'];
  if (headerKey.includes('dieuduong')) return ['dieu_duong'];
  if (headerKey === 'ktv' || headerKey.includes('kythuatvien')) return ['ktv'];
  if (headerKey === 'bs' || headerKey === 'bacsi' || headerKey.includes('ptv') || headerKey.includes('phumo')) {
    return ['bac_si', 'bsnt'];
  }
  return ['bac_si', 'bsnt', 'dieu_duong', 'ktv'];
}

function getStaffSuggestions(text, input = state.focusedInput) {
  const q = String(text || '').trim();
  if (!q) return [];
  const key = normalizeKey(q);
  if (!key) return [];
  const allowedRoles = new Set(allowedStaffRolesForInput(input));
  return state.staff
    .filter(item => item.active !== false && allowedRoles.has(item.vaiTro))
    .map(item => ({ item, score: staffMatchScore(item, key) }))
    .filter(x => x.score < 999)
    .sort((a, b) => a.score - b.score || String(a.item.biDanh).localeCompare(String(b.item.biDanh), 'vi'))
    .slice(0, 8)
    .map(x => x.item);
}

function hideSuggestPanel() {
  const panel = $('suggestPanel');
  if (!panel) return;
  panel.classList.add('hidden');
  panel.innerHTML = '';
  state.suggestItems = [];
  state.suggestActiveIndex = -1;
}

function updateStaffSuggestion(input) {
  if (!input || input.readOnly || !isStaffHeader(input.dataset.header)) {
    hideSuggestPanel();
    return;
  }
  const items = getStaffSuggestions(input.value, input);
  if (!items.length) {
    hideSuggestPanel();
    return;
  }
  state.suggestItems = items;
  state.suggestActiveIndex = 0;
  renderSuggestPanel(input);
}

function renderSuggestPanel(input) {
  const panel = $('suggestPanel');
  if (!panel) return;
  const rect = input.getBoundingClientRect();
  const width = Math.min(Math.max(rect.width + 80, 240), 360);
  const left = Math.min(rect.left, window.innerWidth - width - 12);
  const top = Math.min(rect.bottom + 4, window.innerHeight - 230);
  panel.style.left = `${Math.max(8, left)}px`;
  panel.style.top = `${Math.max(8, top)}px`;
  panel.style.width = `${width}px`;
  panel.innerHTML = state.suggestItems.map((item, idx) => `
    <button type="button" class="suggest-item ${idx === state.suggestActiveIndex ? 'active' : ''}" data-idx="${idx}">
      <span class="suggest-alias">${escapeHtml(item.biDanh)}</span>
      <span class="suggest-name">${escapeHtml(item.hoTen || '')}</span>
      <span class="suggest-group">${escapeHtml(item.nhom || '')}</span>
    </button>
  `).join('');
  panel.classList.remove('hidden');
  panel.querySelectorAll('[data-idx]').forEach(btn => {
    btn.addEventListener('mousedown', event => {
      event.preventDefault();
      chooseSuggestion(Number(btn.dataset.idx));
    });
  });
}

function chooseSuggestion(index = state.suggestActiveIndex) {
  const input = state.focusedInput;
  const item = state.suggestItems[index];
  if (!input || !item) return;
  input.value = item.biDanh;
  input.focus();
  input.dispatchEvent(new Event('input', { bubbles: true }));
  hideSuggestPanel();
}

function moveSuggestion(delta) {
  if (!state.suggestItems.length) return;
  state.suggestActiveIndex = (state.suggestActiveIndex + delta + state.suggestItems.length) % state.suggestItems.length;
  if (state.focusedInput) renderSuggestPanel(state.focusedInput);
}


function staffNamesForInfo() {
  return state.staff.length ? `${state.staff.length} bí danh` : 'chưa tải bí danh';
}

async function loadStaffList() {
  try {
    const data = await api('/api/staff-list');
    state.staff = data.staff || [];
    renderStaffDatalist();
    renderStaffManager();
    log(`Đã tải ${state.staff.length} nhân sự.`);
  } catch (err) {
    log(`Không tải được danh sách nhân sự: ${err.message}`);
  }
}


async function loadClsList() {
  try {
    const data = await api('/api/cls-list');
    state.clsList = data.cls || [];
    renderClsManager();
    log(`Đã tải ${state.clsList.length} tên CLS chuyển tiểu phẫu.`);
  } catch (err) {
    log(`Không tải được danh mục tên CLS: ${err.message}`);
  }
}

function addFileOptionIfMissing(file) {
  const select = $('fileSelect');
  if (!select || !file) return;
  const exists = Array.from(select.options).some(opt => opt.value === file);
  if (!exists) {
    const opt = document.createElement('option');
    opt.value = file;
    opt.textContent = displayPath(file);
    select.insertBefore(opt, select.firstChild);
  }
  select.value = file;
}

async function loadFiles() {
  setBusy(true, 'Đang tải danh sách file Excel...');
  try {
    const data = await api('/api/files');
    const select = $('fileSelect');
    select.innerHTML = '';
    data.files.forEach(file => {
      const opt = document.createElement('option');
      opt.value = file.path;
      opt.textContent = file.relativePath;
      select.appendChild(opt);
    });
    const workFile = data.files.find(f => /_NHAP_LIEU\.(xlsx|xlsm)$/i.test(f.name));
    const sourceFile = data.files.find(f => /PM\s*khoa\s*CTCH/i.test(f.name));
    const preferred = workFile || sourceFile || data.files[0];
    if (preferred) {
      await selectFile(preferred.path);
      await loadSheet();
    }
    log(`Đã tải ${data.files.length} file Excel.`);
  } finally {
    setBusy(false);
  }
}

async function loadTheoSoFiles() {
  const list = $('theoSoFileList');
  if (!list) return;
  try {
    const data = await api('/api/theo-so/files');
    list.innerHTML = '';
    (data.files || []).forEach(file => {
      const li = document.createElement('li');
      li.textContent = file.name;
      list.appendChild(li);
    });
  } catch (err) {
    log(`Không tải được danh sách file TheoSo: ${err.message}`);
  }
}

async function uploadTheoSoFiles() {
  const input = $('theoSoUploadInput');
  if (!input.files.length) {
    alert('Bạn chưa chọn file Excel.');
    return;
  }
  setBusy(true, 'Đang upload file vào TheoSo...');
  try {
    const form = new FormData();
    Array.from(input.files).forEach(file => form.append('files', file));
    const res = await fetch('/api/theo-so/upload', { method: 'POST', body: form });
    const data = await res.json();
    if (!res.ok || data.ok === false) throw new Error(data.error || 'Upload lỗi.');
    input.value = '';
    if ($('theoSoUploadFileName')) $('theoSoUploadFileName').textContent = 'Chọn file Excel';
    await loadTheoSoFiles();
    showToast(`Đã upload ${data.uploaded.length} file vào TheoSo`, 'success', 5200);
    log(`Đã upload vào TheoSo: ${data.uploaded.join(', ')}`);
  } catch (err) {
    alert(err.message);
    log(`Lỗi upload TheoSo: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

async function selectFile(file) {
  if (!file) return;
  const changingFile = Boolean(state.currentFile && file !== state.currentFile);

  // Phải lưu xong file/sheet hiện tại trước khi đổi ngữ cảnh. Nếu lưu lỗi,
  // hàm sẽ ném lỗi và state hiện tại vẫn được giữ nguyên.
  if (
    changingFile &&
    (Object.keys(state.rowChanges || {}).length || state.newRowValues)
  ) {
    await flushAllRows();
  }

  const info = await api(`/api/workbook-info?file=${encodeURIComponent(file)}`);

  if (changingFile) {
    // Nháp của file cũ đã được lưu trong localStorage/máy chủ bởi flushAllRows.
    // Chỉ xóa state sau khi xác nhận file mới mở được thành công.
    state.rowChanges = {};
    state.newRowValues = null;
    state.selectedRowNumber = null;
  }
  const previousSheet = state.currentSheet;
  state.currentFile = info.file || file;
  state.sheets = info.sheets || [];
  addFileOptionIfMissing(state.currentFile);
  $('fileInfo').textContent = displayPath(state.currentFile);
  $('downloadCurrent').href = `/download?finalize=1&file=${encodeURIComponent(state.currentFile)}`;
  $('downloadCurrent').classList.remove('hidden');
  const sheetSelect = $('sheetSelect');
  sheetSelect.innerHTML = '';
  state.sheets.forEach(sheet => {
    const opt = document.createElement('option');
    opt.value = sheet;
    opt.textContent = sheet;
    sheetSelect.appendChild(opt);
  });
  const currentStillExists = previousSheet && state.sheets.includes(previousSheet);
  const preferred = currentStillExists ? previousSheet : (state.sheets.find(s => ['phauthuat','thu thuat','tieuphau'].includes(s.toLowerCase())) || state.sheets[0]);
  if (preferred) sheetSelect.value = preferred;
  state.currentSheet = sheetSelect.value;
  updateWorkspaceIdentity();
  log(`Đã chọn file: ${displayPath(state.currentFile)}`);
}

async function uploadFile() {
  const input = $('uploadInput');
  if (!input.files.length) {
    alert('Bạn chưa chọn file Excel.');
    return;
  }
  setBusy(true, 'Đang upload và chuẩn hóa ngày tháng...');
  try {
    const form = new FormData();
    form.append('file', input.files[0]);
    const res = await fetch('/api/upload', { method: 'POST', body: form });
    const data = await res.json();
    if (!res.ok || data.ok === false) throw new Error(data.error || 'Upload lỗi.');
    await loadFiles();
    await selectFile(data.file);
    await loadSheet();
    state.dismissedSampleNoticeFor = '';
    input.value = '';
    if ($('uploadFileName')) $('uploadFileName').textContent = 'Chọn file Excel';
    updateContextUI();
    const dateFix = data.dateNormalization || {};
    const fixedCount = Number(dateFix.fixedCount || 0);
    const fixedMessage = fixedCount > 0
      ? ` · Đã tự sửa ${fixedCount} ô ngày/giờ`
      : ' · Ngày tháng đã được kiểm tra';
    showToast(`Đã upload và mở ${data.name}${fixedMessage}`, 'success', 5200);
    log(`Đã upload: ${data.name}. ${dateFix.message || 'Đã kiểm tra ngày tháng.'}`);
    if (Number(dateFix.skippedInvalidCount || 0) > 0) {
      log(`Có ${dateFix.skippedInvalidCount} giá trị số ở cột ngày nằm ngoài phạm vi 1990–2100 nên không tự đổi.`);
    }
  } catch (err) {
    alert(err.message);
    log(`Lỗi upload: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

async function loadSheet(targetSheet = null) {
  if (!state.currentFile) return;
  const requestedSheet = targetSheet || $('sheetSelect').value || state.currentSheet;
  const previousSheet = state.currentSheet;
  const query = $('searchInput').value.trim();
  const missingOnly = $('missingOnly').checked ? '1' : '0';
  setBusy(true, `Đang đọc sheet ${requestedSheet}...`);
  try {
    await flushAllRows();
    const fileToRead = state.currentFile;
    const data = await api(`/api/sheet-data?file=${encodeURIComponent(fileToRead)}&sheet=${encodeURIComponent(requestedSheet)}&query=${encodeURIComponent(query)}&missingOnly=${missingOnly}&limit=0`);
    state.currentSheet = data.sheet || requestedSheet;
    $('sheetSelect').value = state.currentSheet;
    state.headers = data.headers || [];
    state.rows = data.rows || [];
    state.sheetMeta = { ...data };
    state.visibleHeaders = chooseDisplayHeaders();
    state.selectedRowNumber = null;
    state.rowChanges = {};
    state.newRowValues = null;
    renderDatalists();
    renderTable();
    updateSelectedInfo();
    requestAnimationFrame(() => focusFirstSurgeryEntry());
    const requiredInfo = (data.requiredHeaders || []).length
      ? `Cột bắt buộc: ${(data.requiredHeaders || []).join(', ')}`
      : '';
    // Số dòng/dòng thiếu đã hiện ở các ô số liệu bên phải (data-overview) —
    // không lặp lại ở đây để đỡ rối, chỉ còn thông tin dòng Tổng Cộng + cột bắt buộc.
    $('sheetSummary').textContent = `Tổng Cộng ở dòng ${data.totalRow}${requiredInfo ? ' · ' + requiredInfo : ''}`;
    updateContextUI();
    updateDashboardMetrics();
    log(`Đã đọc sheet ${data.sheet}: ${data.matchedRowCount ?? data.rowCount} dòng.`);
  } catch (err) {
    state.currentSheet = previousSheet;
    if (previousSheet && state.sheets.includes(previousSheet)) $('sheetSelect').value = previousSheet;
    log(`Lỗi đọc sheet: ${err.message}`);
    throw err;
  } finally {
    setBusy(false);
  }
}

function renderTable() {
  const table = $('dataTable');
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  const headers = state.visibleHeaders;
  thead.innerHTML = '<tr>' + headers.map((h, index) => `<th class="${dataCellClass(h, index)}">${escapeHtml(h)}</th>`).join('') + '</tr>';
  tbody.innerHTML = '';

  if (state.newRowValues) {
    tbody.appendChild(buildNewRowElement());
  }

  state.rows.forEach(row => {
    tbody.appendChild(buildDataRowElement(row));
  });
  updateNewRowButtonVisibility();
  updateDashboardMetrics();
  requestAnimationFrame(resetDocumentHorizontalScroll);
}

function buildNewRowElement() {
  const tr = document.createElement('tr');
  tr.className = 'new-row';
  tr.dataset.rowNumber = 'new';
  tr.title = 'Dòng mới đang nhập, bấm Chèn dòng mới vào Excel để lưu chính thức.';
  tr.innerHTML = state.visibleHeaders.map((header, index) => {
    const readonly = isReadOnlyHeader(header);
    const value = state.newRowValues?.[header] ?? '';
    return `<td class="${dataCellClass(header, index)}">${buildInputHtml(header, value, readonly, 'new')}</td>`;
  }).join('');
  attachRowEvents(tr, null, true);
  return tr;
}

function buildDataRowElement(row) {
  const rowNumber = row.rowNumber;
  const tr = document.createElement('tr');
  tr.dataset.rowNumber = rowNumber;
  if (state.selectedRowNumber === rowNumber) tr.classList.add('selected');
  if ((row.missing || []).length) tr.classList.add('missing');
  if (state.rowChanges[rowNumber]) tr.classList.add('dirty');

  tr.innerHTML = state.visibleHeaders.map((header, index) => {
    const changedValue = state.rowChanges[rowNumber]?.[header];
    const value = changedValue !== undefined ? changedValue : (row.values[header] ?? '');
    const readonly = isReadOnlyHeader(header) || isFormulaLike(value);
    const changed = changedValue !== undefined;
    return `<td class="${dataCellClass(header, index)}">${buildInputHtml(header, value, readonly, rowNumber, changed)}</td>`;
  }).join('');
  attachRowEvents(tr, row, false);
  return tr;
}

function updateNewRowButtonVisibility() {
  const btn = $('commitNewRowBtn');
  if (!btn) return;
  btn.classList.toggle('hidden', !state.newRowValues);
}


function buildInputHtml(header, value, readonly, rowNumber, changed = false) {
  const staffAttr = isStaffHeader(header) ? ' data-staff-field="1" autocomplete="off"' : '';
  const readonlyAttr = readonly ? ' readonly aria-readonly="true" tabindex="-1"' : ' tabindex="0"';
  const editLabel = !readonly && sheetKind() === 'surgery' ? ` aria-label="Nhập ${escapeHtml(header)}"` : '';
  return `<input class="${cellInputClass(header, changed, readonly)}" value="${escapeHtml(value)}" data-row="${escapeHtml(rowNumber)}" data-header="${escapeHtml(header)}"${staffAttr}${readonlyAttr}${editLabel} />`;
}


function attachRowEvents(tr, row, isNew) {
  tr.querySelectorAll('.cell-input').forEach(input => {
    input.addEventListener('focus', () => {
      state.focusedInput = input;
      hideSuggestPanel();
      document.querySelectorAll('#dataTable .cell-input.active-entry').forEach(el => el.classList.remove('active-entry'));
      if (!input.readOnly) input.classList.add('active-entry');
      if (!isNew) selectRow(Number(input.dataset.row), false);
    });
    input.addEventListener('click', () => {
      state.focusedInput = input;
      hideSuggestPanel();
      if (!isNew) selectRow(Number(input.dataset.row), false);
    });
    input.addEventListener('input', () => {
      if (input.readOnly) return;
      if (isNew) handleNewRowInput(input);
      else handleCellInput(input);
      updateStaffSuggestion(input);
    });
    input.addEventListener('change', () => {
      if (input.readOnly) return;
      if (isNew) handleNewRowInput(input);
      else handleCellInput(input);
    });
    input.addEventListener('blur', () => {
      setTimeout(() => {
        if (document.activeElement !== input) hideSuggestPanel();
      }, 120);
    });
    input.addEventListener('keydown', handleCellKeydown);
  });
}


function selectRow(rowNumber, scroll = false) {
  state.selectedRowNumber = Number(rowNumber);
  updateSelectedInfo();
  document.querySelectorAll('#dataTable tbody tr').forEach(tr => {
    tr.classList.toggle('selected', Number(tr.dataset.rowNumber) === state.selectedRowNumber);
  });
  if (scroll) {
    const tr = document.querySelector(`#dataTable tbody tr[data-row-number="${state.selectedRowNumber}"]`);
    if (tr) tr.scrollIntoView({ block: 'nearest' });
  }
}

function updateSelectedInfo() {
  const el = $('selectedInfo');
  if (!el) return;
  if (!state.selectedRowNumber) {
    el.textContent = 'Chưa chọn';
  } else {
    const dirty = state.rowChanges[state.selectedRowNumber] ? ' · có thay đổi chưa lưu xong' : '';
    el.textContent = `Dòng Excel ${state.selectedRowNumber}${dirty}`;
  }
}

function handleCellInput(input) {
  const rowNumber = Number(input.dataset.row);
  const header = input.dataset.header;
  const value = input.value;
  const row = state.rows.find(r => Number(r.rowNumber) === rowNumber);
  if (!row) return;

  if (!state.rowChanges[rowNumber]) state.rowChanges[rowNumber] = {};
  const original = row.values[header] ?? '';
  if (String(value) === String(original ?? '')) {
    delete state.rowChanges[rowNumber][header];
    if (!Object.keys(state.rowChanges[rowNumber]).length) delete state.rowChanges[rowNumber];
    input.classList.remove('changed');
  } else {
    state.rowChanges[rowNumber][header] = value;
    input.classList.add('changed');
  }
  const tr = input.closest('tr');
  if (tr) tr.classList.toggle('dirty', Boolean(state.rowChanges[rowNumber]));
  updateSelectedInfo();
  updateDirtyMetrics();
  scheduleDraftSave();
  scheduleRowAutosave(rowNumber);
}

function handleNewRowInput(input) {
  const header = input.dataset.header;
  if (!state.newRowValues) state.newRowValues = {};
  state.newRowValues[header] = input.value;
  updateDirtyMetrics();
  input.classList.add('changed');
  scheduleDraftSave();
  autosaveStatus('Đã lưu nháp dòng mới · bấm Chèn dòng để ghi Excel', 'draft');
}

function handleCellKeydown(event) {
  const input = event.target;
  const panelOpen = !$('suggestPanel')?.classList.contains('hidden') && state.suggestItems.length;
  if (panelOpen) {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      moveSuggestion(1);
      return;
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault();
      moveSuggestion(-1);
      return;
    }
    if (event.key === 'Enter' || event.key === 'Tab') {
      event.preventDefault();
      chooseSuggestion(state.suggestActiveIndex);
      // Sau khi chấp nhận bí danh, chuyển thẳng sang ô cần nhập kế tiếp.
      setTimeout(() => focusSequentialEditable(input, event.shiftKey ? -1 : 1), 0);
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      hideSuggestPanel();
      return;
    }
  }

  if (event.key === 'Tab') {
    event.preventDefault();
    hideSuggestPanel();
    focusSequentialEditable(input, event.shiftKey ? -1 : 1);
    return;
  }

  if (event.key === 'Enter') {
    event.preventDefault();
    hideSuggestPanel();
    focusVerticalEditable(input, event.shiftKey ? -1 : 1);
    return;
  }

  if (event.key === 'ArrowDown') {
    event.preventDefault();
    hideSuggestPanel();
    focusVerticalEditable(input, 1);
    return;
  }

  if (event.key === 'ArrowUp') {
    event.preventDefault();
    hideSuggestPanel();
    focusVerticalEditable(input, -1);
    return;
  }

  if (event.key === 'ArrowRight') {
    const atEnd = input.selectionStart === input.value.length && input.selectionEnd === input.value.length;
    if (event.ctrlKey || event.altKey || atEnd) {
      event.preventDefault();
      hideSuggestPanel();
      focusSequentialEditable(input, 1);
    }
    return;
  }

  if (event.key === 'ArrowLeft') {
    const atStart = input.selectionStart === 0 && input.selectionEnd === 0;
    if (event.ctrlKey || event.altKey || atStart) {
      event.preventDefault();
      hideSuggestPanel();
      focusSequentialEditable(input, -1);
    }
  }
}

function editableInputs() {
  return Array.from(document.querySelectorAll('#dataTable tbody .cell-input:not([readonly])'))
    .filter(input => input.offsetParent !== null);
}

function focusEditableInput(target) {
  if (!target) return false;
  target.focus({ preventScroll: true });
  if (typeof target.select === 'function') target.select();
  target.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'smooth' });
  return true;
}

function focusFirstSurgeryEntry() {
  if (sheetKind() !== 'surgery') return false;
  const ptvHeader = findHeader(['PTV chính']);
  const candidates = editableInputs();
  const firstMissingPtv = ptvHeader
    ? candidates.find(input => input.dataset.header === ptvHeader && !String(input.value || '').trim())
    : null;
  return focusEditableInput(firstMissingPtv || candidates[0]);
}

function focusSequentialEditable(input, delta) {
  const cells = editableInputs();
  const currentIndex = cells.indexOf(input);
  if (currentIndex < 0) return false;
  const target = cells[currentIndex + delta];
  if (target) return focusEditableInput(target);
  showToast(delta > 0 ? 'Đã đến ô nhập cuối cùng.' : 'Đã ở ô nhập đầu tiên.', 'info', 1800);
  return false;
}

function focusVerticalEditable(input, rowDelta) {
  const currentHeader = input.dataset.header;
  const rows = Array.from(document.querySelectorAll('#dataTable tbody tr'))
    .filter(row => row.offsetParent !== null);
  const currentRow = input.closest('tr');
  const rowIndex = rows.indexOf(currentRow);
  if (rowIndex < 0) return false;

  for (let index = rowIndex + rowDelta; index >= 0 && index < rows.length; index += rowDelta) {
    const target = rows[index].querySelector(`.cell-input[data-header="${cssEscape(currentHeader)}"]:not([readonly])`);
    if (target) return focusEditableInput(target);
  }
  showToast(rowDelta > 0 ? 'Đã đến dòng cuối cùng.' : 'Đã ở dòng đầu tiên.', 'info', 1800);
  return false;
}

function focusRelativeCell(input, rowDelta, colDelta) {
  if (rowDelta) return focusVerticalEditable(input, rowDelta);
  if (colDelta) return focusSequentialEditable(input, colDelta);
  return false;
}

function cssEscape(value) {
  if (window.CSS && CSS.escape) return CSS.escape(value);
  return String(value).replace(/"/g, '\\"');
}

function currentDraftPayload() {
  if (!state.currentFile || !state.currentSheet) return null;
  return {
    file: state.currentFile,
    sheet: state.currentSheet,
    selectedRowNumber: state.selectedRowNumber,
    rowChanges: state.rowChanges || {},
    newRowValues: state.newRowValues || null,
    viewMode: $('viewModeSelect') ? $('viewModeSelect').value : 'auto',
    updatedAt: new Date().toISOString(),
  };
}

function saveDraftLocal(draft) {
  if (!draft) return;
  try { localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(draft)); } catch (_err) {}
}

async function saveDraftServer(draft) {
  if (!draft) return;
  await api('/api/autosave-state', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(draft)
  });
}

function scheduleDraftSave() {
  const draft = currentDraftPayload();
  if (!draft) return;
  const signature = JSON.stringify(draft);
  if (signature === state.lastDraftSignature) return;
  state.lastDraftSignature = signature;
  clearTimeout(state.autosaveDraftTimer);
  state.autosaveDraftTimer = setTimeout(() => {
    saveDraftLocal(draft);
    saveDraftServer(draft).catch(err => log(`Không lưu được nháp máy chủ: ${err.message}`));
  }, 700);
}

function getLocalDraft() {
  try {
    const raw = localStorage.getItem(AUTOSAVE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (_err) {
    return null;
  }
}

async function getSavedDraftCandidates() {
  const localDraft = getLocalDraft();
  let serverDraft = null;
  try {
    const data = await api('/api/autosave-state');
    serverDraft = data.state || null;
  } catch (_err) {}
  const candidates = [localDraft, serverDraft].filter(d => d && d.file && d.sheet);
  candidates.sort((a, b) => new Date(b.updatedAt || 0) - new Date(a.updatedAt || 0));
  return candidates;
}
async function persistCleanState(message = '') {
  const clean = {
    file: state.currentFile,
    sheet: state.currentSheet,
    selectedRowNumber: state.selectedRowNumber || null,
    rowChanges: {},
    newRowValues: null,
    viewMode: $('viewModeSelect') ? $('viewModeSelect').value : 'auto',
    updatedAt: new Date().toISOString(),
  };
  state.rowChanges = {};
  state.newRowValues = null;
  state.lastDraftSignature = JSON.stringify(clean);
  saveDraftLocal(clean);
  try {
    await saveDraftServer(clean);
  } catch (err) {
    log(`Đã lưu file nhưng không cập nhật được nháp máy chủ: ${err.message}`);
  }
  if (message) autosaveStatus(message, 'ok');
  updateNewRowButtonVisibility();
}


async function restoreDraft(draft, silent = false) {
  if (!draft || !draft.file) return false;
  try {
    if ($('viewModeSelect') && draft.viewMode) $('viewModeSelect').value = draft.viewMode;
    await selectFile(draft.file);
    const targetSheet = state.sheets.includes(draft.sheet) ? draft.sheet : state.currentSheet;
    await loadSheet(targetSheet);
    state.rowChanges = sanitizeChangesForCurrentSheet(draft.rowChanges || {});
    // Sheet phẫu thuật chỉ nhập trên các dòng có sẵn, không khôi phục dòng mới.
    state.newRowValues = sheetKind() === 'surgery' ? null : (draft.newRowValues || null);
    state.selectedRowNumber = draft.selectedRowNumber || null;
    applyDraftChangesToRows();
    renderTable();
    updateSelectedInfo();
    autosaveStatus(`Đã khôi phục bản nháp ${new Date(draft.updatedAt || Date.now()).toLocaleTimeString('vi-VN')}`, 'ok');
    if (!silent) log('Đã khôi phục dữ liệu tự lưu gần nhất.');
    const dirtyRows = Object.keys(state.rowChanges || {});
    dirtyRows.forEach(rowNumber => scheduleRowAutosave(Number(rowNumber), 400));
    return true;
  } catch (err) {
    autosaveStatus('Không khôi phục được nháp', 'error');
    log(`Không khôi phục được nháp: ${err.message}`);
    return false;
  }
}

function applyDraftChangesToRows() {
  Object.entries(state.rowChanges || {}).forEach(([rowNumber, changes]) => {
    const row = state.rows.find(r => Number(r.rowNumber) === Number(rowNumber));
    if (!row) return;
    Object.entries(changes || {}).forEach(([header, value]) => {
      row.values[header] = value;
    });
  });
}

async function restoreBestDraftIfAny(silent = true) {
  const candidates = await getSavedDraftCandidates();
  for (const draft of candidates) {
    if (await restoreDraft(draft, silent)) return true;
  }
  return false;
}

async function clearAutosaveDraft() {
  try { localStorage.removeItem(AUTOSAVE_KEY); } catch (_err) {}
  try { await api('/api/clear-autosave-state', { method: 'POST' }); } catch (_err) {}
  state.rowChanges = {};
  state.newRowValues = null;
  state.lastDraftSignature = '';
  autosaveStatus('Đã xóa nháp tự lưu', 'ok');
  updateDirtyMetrics();
  renderTable();
  log('Đã xóa nháp tự lưu.');
}

function currentSaveContext() {
  return {
    file: state.currentFile,
    sheet: state.currentSheet,
  };
}

function scheduleRowAutosave(_rowNumber, delay = 1200) {
  clearTimeout(state.batchSaveTimer);
  autosaveStatus('Đang chờ tự động lưu...', 'saving');
  state.batchSaveTimer = setTimeout(() => {
    enqueueBatchSave(currentSaveContext()).catch(err => {
      autosaveStatus('Tự động lưu lỗi · dữ liệu vẫn được giữ trong nháp', 'error');
      log(`Lỗi tự động lưu: ${err.message}`);
    });
  }, delay);
}

function enqueueBatchSave(target = currentSaveContext()) {
  const run = state.saveQueue.then(() => saveDirtyRowsBatch(target));
  // Giữ hàng đợi tiếp tục hoạt động sau lỗi, nhưng promise trả về cho caller vẫn
  // giữ trạng thái reject để không che giấu lỗi.
  state.saveQueue = run.catch(() => {});
  return run;
}

async function saveDirtyRowsBatch(target) {
  if (!target?.file || !target?.sheet) return null;
  if (target.file !== state.currentFile || target.sheet !== state.currentSheet) {
    throw new Error('Ngữ cảnh file/sheet đã thay đổi trước khi lưu. Dữ liệu nháp chưa bị xóa.');
  }

  const snapshot = {};
  const safeChanges = sanitizeChangesForCurrentSheet(state.rowChanges || {});
  // Đồng bộ lại state để bản nháp cũ không thể mang theo thay đổi ở cột bị khóa.
  state.rowChanges = safeChanges;
  Object.entries(safeChanges).forEach(([rowNumber, changes]) => {
    if (changes && Object.keys(changes).length) snapshot[rowNumber] = { ...changes };
  });
  const rowNumbers = Object.keys(snapshot).map(Number).filter(Boolean);
  if (!rowNumbers.length) return null;

  rowNumbers.forEach(rowNumber => {
    const tr = document.querySelector(`#dataTable tbody tr[data-row-number="${rowNumber}"]`);
    if (tr) tr.classList.add('saving');
  });
  autosaveStatus(`Đang lưu ${rowNumbers.length} dòng trong một lần...`, 'saving');

  let data;
  try {
    data = await api('/api/update-rows', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file: target.file,
        sheet: target.sheet,
        rows: rowNumbers.map(rowNumber => ({ rowNumber, data: snapshot[rowNumber] })),
        overwrite: false,
        suffix: '_NHAP_LIEU'
      })
    });
  } catch (err) {
    rowNumbers.forEach(rowNumber => {
      const tr = document.querySelector(`#dataTable tbody tr[data-row-number="${rowNumber}"]`);
      if (tr) tr.classList.remove('saving');
    });
    throw err;
  }

  // Chỉ cập nhật state khi người dùng vẫn đang ở đúng sheet đã lưu.
  if (state.currentFile === target.file && state.currentSheet === target.sheet) {
    state.currentFile = data.file;
    addFileOptionIfMissing(data.file);
    $('fileInfo').textContent = displayPath(data.file);
    $('downloadCurrent').href = `/download?finalize=1&file=${encodeURIComponent(data.file)}`;
    $('downloadCurrent').classList.remove('hidden');

    rowNumbers.forEach(rowNumber => {
      const savedValues = snapshot[rowNumber];
      const row = state.rows.find(r => Number(r.rowNumber) === Number(rowNumber));
      if (row) Object.assign(row.values, savedValues);

      const latest = state.rowChanges[rowNumber] || {};
      Object.entries(savedValues).forEach(([header, value]) => {
        if (String(latest[header] ?? '') === String(value ?? '')) delete latest[header];
      });
      if (!Object.keys(latest).length) delete state.rowChanges[rowNumber];
      else state.rowChanges[rowNumber] = latest;

      const tr = document.querySelector(`#dataTable tbody tr[data-row-number="${rowNumber}"]`);
      if (tr) {
        tr.classList.remove('saving');
        tr.classList.toggle('dirty', Boolean(state.rowChanges[rowNumber]));
        tr.classList.add('saved-flash');
        tr.querySelectorAll('.cell-input').forEach(input => {
          if (savedValues[input.dataset.header] !== undefined && !state.rowChanges[rowNumber]?.[input.dataset.header]) {
            input.classList.remove('changed');
          }
        });
        setTimeout(() => tr.classList.remove('saved-flash'), 1500);
      }
    });
  }

  if ((data.warnings || []).length) {
    (data.warnings || []).forEach(item => log(`Cảnh báo khi lưu: ${item}`));
  }
  saveDraftLocal(currentDraftPayload());
  scheduleDraftSave();
  updateSelectedInfo();
  updateDirtyMetrics();
  autosaveStatus(`Đã lưu ${rowNumbers.length} dòng lúc ${new Date().toLocaleTimeString('vi-VN')}`, 'ok');
  return data;
}

async function flushAllRows() {
  clearTimeout(state.autosaveDraftTimer);
  clearTimeout(state.batchSaveTimer);
  state.batchSaveTimer = null;

  const draft = currentDraftPayload();
  saveDraftLocal(draft);
  // Lỗi lưu nháp máy chủ không được làm mất dữ liệu; localStorage vẫn giữ nháp.
  try {
    await saveDraftServer(draft);
  } catch (err) {
    log(`Không lưu được nháp máy chủ: ${err.message}`);
  }

  await state.saveQueue;
  if (Object.keys(state.rowChanges || {}).length) {
    await enqueueBatchSave(currentSaveContext());
  }
  return true;
}

async function saveAllRows() {
  const rows = Object.keys(state.rowChanges || {}).map(Number).filter(Boolean);
  if (!rows.length) {
    autosaveStatus('Không có dòng nào cần lưu', 'ok');
    return;
  }
  setBusy(true, 'Đang lưu các dòng đang sửa...');
  try {
    await flushAllRows();
    autosaveStatus('Đã lưu tất cả dòng đang sửa', 'ok');
    log('Đã lưu tất cả dòng đang sửa.');
  } catch (err) {
    autosaveStatus('Lưu thất bại · dữ liệu vẫn còn trong nháp', 'error');
    alert(err.message);
    log(`Lỗi lưu tất cả: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

function startNewRow(values = null) {
  if (sheetKind() === 'surgery') {
    showToast('Sheet phẫu thuật chỉ cho nhập PTV chính và Phụ mổ 1–3 trên các dòng có sẵn.', 'info', 3200);
    return;
  }
  state.newRowValues = {};
  updateDirtyMetrics();
  state.visibleHeaders.forEach(header => {
    if (!isReadOnlyHeader(header)) state.newRowValues[header] = values?.[header] ?? '';
  });
  renderTable();
  autosaveStatus('Đang nhập dòng mới · chưa chèn vào Excel', 'draft');
  saveDraftLocal(currentDraftPayload());
  scheduleDraftSave();
  const firstInput = document.querySelector('#dataTable tbody tr.new-row .cell-input:not([readonly])');
  if (firstInput) firstInput.focus();
}

function copySelectedRow() {
  if (!state.selectedRowNumber) {
    alert('Chưa chọn dòng để copy.');
    return;
  }
  const row = state.rows.find(r => Number(r.rowNumber) === Number(state.selectedRowNumber));
  if (!row) return;
  startNewRow(row.values || {});
  log(`Đã copy dòng ${state.selectedRowNumber} thành dòng mới. Kiểm tra rồi bấm Chèn dòng.`);
}

async function insertNewRow() {
  if (!state.newRowValues) {
    alert('Chưa có dòng mới để chèn.');
    return;
  }
  setBusy(true, 'Đang thêm dòng trước Tổng Cộng...');
  try {
    await flushAllRows();
    const data = await api('/api/insert-row', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file: state.currentFile,
        sheet: state.currentSheet,
        data: state.newRowValues || {},
        overwrite: false,
        suffix: '_NHAP_LIEU'
      })
    });
    state.currentFile = data.file;
    state.newRowValues = null;
    await clearAutosaveDraft();
    autosaveStatus('Đã thêm dòng và lưu vào Excel', 'ok');
    showResult($('taskResult'), data);
    log(`Đã thêm dòng ${data.rowNumber}.`);
    await selectFile(data.file);
    $('sheetSelect').value = state.currentSheet;
    await loadSheet();
    await persistCleanState('Đã thêm dòng và lưu vào file làm việc');
  } catch (err) {
    alert(err.message);
    log(`Lỗi thêm dòng: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

async function deleteSelectedRow() {
  if (!state.selectedRowNumber) {
    alert('Chưa chọn dòng để xóa.');
    return;
  }
  const ok = confirm(`Xóa dòng Excel ${state.selectedRowNumber}?`);
  if (!ok) return;
  setBusy(true, 'Đang xóa dòng...');
  try {
    await flushAllRows();
    const data = await api('/api/delete-row', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file: state.currentFile,
        sheet: state.currentSheet,
        rowNumber: state.selectedRowNumber,
        overwrite: false,
        suffix: '_NHAP_LIEU'
      })
    });
    state.currentFile = data.file;
    state.selectedRowNumber = null;
    autosaveStatus('Đã xóa dòng và lưu vào Excel', 'ok');
    showResult($('taskResult'), data);
    log(`Đã xóa dòng ${data.deletedRow}.`);
    await selectFile(data.file);
    $('sheetSelect').value = state.currentSheet;
    await loadSheet();
    await persistCleanState('Đã xóa dòng và lưu vào file làm việc');
  } catch (err) {
    alert(err.message);
    log(`Lỗi xóa dòng: ${err.message}`);
  } finally {
    setBusy(false);
  }
}



function selectedClearColumns() {
  return Array.from(document.querySelectorAll('.clear-column-check:checked')).map(input => input.value);
}

function clearCountElementId(column) {
  const key = normalizeKey(column);
  if (key === 'phumo1') return 'countPhuMo1';
  if (key === 'phumo2') return 'countPhuMo2';
  if (key === 'phumo3') return 'countPhuMo3';
  if (key === 'ptvchinh') return 'countPtvChinh';
  return '';
}

function updateClearPreviewSummary() {
  const selected = selectedClearColumns();
  const stats = state.clearPreviewData?.columnCounts || {};
  const total = selected.reduce((sum, column) => sum + Number(stats[column] || 0), 0);
  const preview = $('clearPreview');
  if (preview) {
    preview.textContent = selected.length
      ? `Sẽ xóa ${total} ô có dữ liệu trong ${selected.length} cột đã chọn.`
      : 'Chưa chọn cột nào để xóa.';
  }
  const canConfirm = selected.length > 0 && total > 0 && Boolean($('clearConfirmCheck')?.checked);
  if ($('confirmClearAssistantsBtn')) $('confirmClearAssistantsBtn').disabled = !canConfirm;
}

function renderClearColumnCounts(data) {
  const stats = data?.columnCounts || {};
  ['Phụ mổ 1', 'Phụ mổ 2', 'Phụ mổ 3', 'PTV chính'].forEach(column => {
    const id = clearCountElementId(column);
    const el = id ? $(id) : null;
    if (!el) return;
    const count = Number(stats[column] || 0);
    el.textContent = data?.availableColumns?.includes(column)
      ? `${count} ô đang có dữ liệu`
      : 'Không tìm thấy cột trong sheet này';
    const checkbox = document.querySelector(`.clear-column-check[value="${column}"]`);
    if (checkbox) {
      checkbox.disabled = !data?.availableColumns?.includes(column);
      if (checkbox.disabled) checkbox.checked = false;
    }
  });
  updateClearPreviewSummary();
}

async function openClearAssistantsModal() {
  if (!state.currentFile || !state.currentSheet) {
    showToast('Chưa chọn file và sheet cần làm sạch.', 'warning');
    return;
  }
  if (sheetKind() !== 'surgery') {
    showToast('Chức năng này chỉ dùng cho sheet phẫu thuật.', 'warning');
    return;
  }
  const modal = $('clearAssistantsModal');
  modal.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
  $('clearConfirmCheck').checked = false;
  document.querySelectorAll('.clear-column-check').forEach(input => {
    input.checked = normalizeKey(input.value).startsWith('phumo');
    input.disabled = false;
  });
  $('confirmClearAssistantsBtn').disabled = true;
  state.clearPreviewData = null;
  ['countPhuMo1', 'countPhuMo2', 'countPhuMo3', 'countPtvChinh'].forEach(id => {
    if ($(id)) $(id).textContent = 'Đang kiểm tra...';
  });
  try {
    await flushAllRows();
    const data = await api('/api/clear-columns', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file: state.currentFile,
        sheet: state.currentSheet,
        columns: ['Phụ mổ 1', 'Phụ mổ 2', 'Phụ mổ 3', 'PTV chính'],
        preview: true,
      }),
    });
    state.clearPreviewData = data;
    renderClearColumnCounts(data);
  } catch (err) {
    closeClearAssistantsModal();
    showToast(err.message, 'error');
    log(`Không kiểm tra được dữ liệu phụ mổ: ${err.message}`);
  }
}

function closeClearAssistantsModal() {
  $('clearAssistantsModal')?.classList.add('hidden');
  document.body.style.overflow = '';
  state.clearPreviewData = null;
}

async function confirmClearAssistants() {
  const columns = selectedClearColumns();
  if (!columns.length || !$('clearConfirmCheck')?.checked) return;
  setBusy(true, 'Đang xóa dữ liệu mẫu phụ mổ...');
  try {
    const data = await api('/api/clear-columns', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file: state.currentFile,
        sheet: state.currentSheet,
        columns,
        preview: false,
        overwrite: false,
        suffix: '_NHAP_LIEU',
      }),
    });
    closeClearAssistantsModal();
    state.currentFile = data.file || state.currentFile;
    showResult($('taskResult'), {
      ...data,
      message: `Đã xóa ${data.clearedCells || 0} ô dữ liệu mẫu trong ${data.clearedRows || 0} dòng.`,
    });
    await selectFile(state.currentFile);
    if (state.sheets.includes(data.sheet)) $('sheetSelect').value = data.sheet;
    await loadSheet(data.sheet);
    await persistCleanState('Đã xóa dữ liệu mẫu phụ mổ');
    showToast(`Đã xóa ${data.clearedCells || 0} ô dữ liệu mẫu. File gốc vẫn được giữ nguyên.`, 'success', 5000);
    log(`Đã xóa dữ liệu mẫu: ${(data.columns || []).join(', ')}.`);
  } catch (err) {
    showToast(err.message, 'error', 5000);
    log(`Lỗi xóa dữ liệu mẫu phụ mổ: ${err.message}`);
  } finally {
    setBusy(false);
  }
}


async function loadEmrConfig() {
  const data = await api('/api/emr-config');
  state.emrConfig = data.config || {};
  return state.emrConfig;
}

function closeEmrConfigModal() {
  $('emrConfigModal')?.classList.add('hidden');
  document.body.style.overflow = '';
}

async function openEmrConfigModal(pendingAction = '') {
  state.pendingEmrAction = pendingAction || state.pendingEmrAction || '';
  try {
    const config = await loadEmrConfig();
    $('emrUrlInput').value = config.url_login || 'http://192.168.2.26:2026/login.aspx';
    $('emrUsernameInput').value = config.username || '';
    $('emrPasswordInput').value = '';
    $('emrHeadlessInput').checked = config.headless !== false;
  } catch (err) {
    showToast(err.message, 'error', 5000);
  }
  $('emrConfigModal')?.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
  setTimeout(() => $('emrUsernameInput')?.focus(), 60);
}

async function saveEmrConfig() {
  const payload = {
    url_login: $('emrUrlInput').value.trim(),
    username: $('emrUsernameInput').value.trim(),
    password: $('emrPasswordInput').value,
    headless: $('emrHeadlessInput').checked,
  };
  if (!payload.url_login || !payload.username) {
    showToast('Vui lòng nhập địa chỉ và tên đăng nhập EMR.', 'warning', 4500);
    return;
  }
  setBusy(true, 'Đang lưu cấu hình EMR...');
  try {
    const data = await api('/api/emr-config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    state.emrConfig = data.config || {};
    closeEmrConfigModal();
    showToast('Đã lưu cấu hình EMR.', 'success');
    const pending = state.pendingEmrAction;
    state.pendingEmrAction = '';
    if (pending === 'selected') setTimeout(() => runEmrFill(true), 80);
    if (pending === 'all') setTimeout(() => runEmrFill(false), 80);
    if (pending === 'refresh-cls') setTimeout(() => updateMissingClsAndDoctors(), 80);
  } catch (err) {
    showToast(err.message, 'error', 5000);
  } finally {
    setBusy(false);
  }
}

async function updateMissingClsAndDoctors() {
  if (!state.currentFile || sheetSubtype() !== 'tieuphau') {
    showToast('Hãy mở sheet tieuphau trước khi cập nhật CLS còn sót.', 'warning', 4500);
    return;
  }

  try {
    const config = state.emrConfig || await loadEmrConfig();
    if (!config.configured) {
      await openEmrConfigModal('refresh-cls');
      return;
    }
  } catch (_err) {
    await openEmrConfigModal('refresh-cls');
    return;
  }

  let moveData = null;
  setBusy(true, 'Đang chuyển CLS còn sót sang tiểu phẫu...');
  autosaveStatus('Đang cập nhật CLS còn sót...', 'saving');
  try {
    await flushAllRows();

    // Chuyển và lưu workbook trước. Nếu EMR bị gián đoạn, các dòng mới vẫn còn trong file.
    moveData = await api('/api/run-task', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file: state.currentFile, task: 'cap-nhat-cls-tieuphau' })
    });
    state.currentFile = moveData.file || moveData.file_dau_ra || state.currentFile;
    log(`CLS bổ sung: đã chuyển và lưu ${moveData.so_dong_chuyen || 0} dòng.`);

    await selectFile(state.currentFile);
    if (state.sheets.includes('tieuphau')) $('sheetSelect').value = 'tieuphau';
    await loadSheet('tieuphau');
    await persistCleanState(`Đã lưu ${moveData.so_dong_chuyen || 0} dòng CLS còn sót`);

    // EMR chỉ xử lý Bác sĩ trống/không hợp lệ và tự lưu ngay sau từng dòng tìm được.
    setBusy(true, 'Đã lưu CLS; đang lấy Bác sĩ từ EMR...');
    autosaveStatus('Đang lấy Bác sĩ từ EMR...', 'saving');
    const emrData = await api('/api/emr-fill-doctor', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file: state.currentFile, sheet: 'tieuphau', rowNumber: null }),
    });

    state.currentFile = emrData.file || state.currentFile;
    const combined = {
      ...emrData,
      file: state.currentFile,
      so_dong_chuyen: moveData.so_dong_chuyen || 0,
      message: `Đã chuyển ${moveData.so_dong_chuyen || 0} dòng CLS còn sót; EMR đã điền ${emrData.updatedRows || 0} dòng Bác sĩ.`,
    };
    showResult($('taskResult'), combined);

    await selectFile(state.currentFile);
    if (state.sheets.includes('tieuphau')) $('sheetSelect').value = 'tieuphau';
    await loadSheet('tieuphau');
    await persistCleanState('Đã cập nhật CLS còn sót và Bác sĩ');

    const unmatched = emrData.unmatchedRows || 0;
    showToast(
      `Đã thêm ${moveData.so_dong_chuyen || 0} dòng; điền ${emrData.updatedRows || 0} Bác sĩ${unmatched ? `; ${unmatched} dòng cần kiểm tra` : ''}.`,
      unmatched ? 'warning' : 'success',
      7000,
    );
    log(`Hoàn tất cập nhật CLS + EMR: chuyển ${moveData.so_dong_chuyen || 0}, điền BS ${emrData.updatedRows || 0}.`);
  } catch (err) {
    const moved = moveData?.so_dong_chuyen || 0;
    const message = err.message || String(err);
    if (message.includes('Chưa cấu hình tài khoản EMR')) {
      state.emrConfig = null;
      await openEmrConfigModal('refresh-cls');
    } else if (moveData) {
      showResult($('taskResult'), {
        ...moveData,
        message: `Đã chuyển và lưu ${moved} dòng CLS. Phần lấy Bác sĩ từ EMR chưa hoàn tất: ${message}`,
      });
      showToast(`Đã lưu ${moved} dòng CLS; lỗi khi lấy Bác sĩ: ${message}`, 'warning', 8000);
      log(`CLS đã lưu nhưng EMR lỗi: ${message}`);
      try {
        await selectFile(state.currentFile);
        if (state.sheets.includes('tieuphau')) $('sheetSelect').value = 'tieuphau';
        await loadSheet('tieuphau');
      } catch (_reloadErr) {}
    } else {
      showToast(message, 'error', 6500);
      log(`Lỗi cập nhật CLS còn sót: ${message}`);
    }
  } finally {
    autosaveStatus('Tự động lưu đang bật', 'ok');
    setBusy(false);
  }
}

async function runEmrFill(selectedOnly = false) {
  if (!state.currentFile || sheetSubtype() !== 'tieuphau') {
    showToast('Chức năng này chỉ dùng tại sheet tieuphau.', 'warning', 4500);
    return;
  }
  if (selectedOnly && !state.selectedRowNumber) {
    showToast('Hãy bấm chọn một dòng trong bảng trước.', 'warning', 4000);
    return;
  }
  try {
    const config = state.emrConfig || await loadEmrConfig();
    if (!config.configured) {
      await openEmrConfigModal(selectedOnly ? 'selected' : 'all');
      return;
    }
  } catch (_err) {
    await openEmrConfigModal(selectedOnly ? 'selected' : 'all');
    return;
  }

  const label = selectedOnly ? `dòng ${state.selectedRowNumber}` : 'các dòng Bác sĩ trống hoặc không hợp lệ';
  setBusy(true, `Đang đối chiếu EMR cho ${label}...`);
  autosaveStatus('Đang lấy dữ liệu EMR...', 'saving');
  try {
    await flushAllRows();
    const data = await api('/api/emr-fill-doctor', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file: state.currentFile,
        sheet: state.currentSheet,
        rowNumber: selectedOnly ? state.selectedRowNumber : null,
      }),
    });
    state.currentFile = data.file || state.currentFile;
    showResult($('taskResult'), data);
    await selectFile(state.currentFile);
    if (state.sheets.includes('tieuphau')) $('sheetSelect').value = 'tieuphau';
    await loadSheet('tieuphau');
    await persistCleanState('Đã cập nhật Bác sĩ từ EMR');
    showToast(`Đã điền ${data.updatedRows || 0} dòng; ${data.unmatchedRows || 0} dòng cần kiểm tra.`, data.unmatchedRows ? 'warning' : 'success', 6000);
    log(`EMR: cập nhật ${data.updatedRows || 0}/${data.checkedRows || 0} dòng.`);
  } catch (err) {
    const message = err.message || String(err);
    if (message.includes('Chưa cấu hình tài khoản EMR')) {
      state.emrConfig = null;
      await openEmrConfigModal(selectedOnly ? 'selected' : 'all');
    } else {
      showToast(message, 'error', 6500);
      log(`Lỗi lấy Bác sĩ từ EMR: ${message}`);
    }
  } finally {
    autosaveStatus('Tự động lưu đang bật', 'ok');
    setBusy(false);
  }
}

async function runTask(task) {
  if (!state.currentFile) return;
  const assistantHeaders = assistantHeaderMap();
  const sampleCount = [assistantHeaders.phuMo1, assistantHeaders.phuMo2, assistantHeaders.phuMo3]
    .filter(Boolean)
    .reduce((sum, header) => sum + countNonBlankInHeader(header), 0);
  const noticeKey = `${state.currentFile}|${state.currentSheet}`;
  const sourceHasUnreviewedSamples = sheetKind() === 'surgery'
    && !/_NHAP_LIEU\.(xlsx|xlsm)$/i.test(state.currentFile)
    && sampleCount > 0
    && state.dismissedSampleNoticeFor !== noticeKey;
  if ((task === 'bo-sung-phu-mo' || task === 'tat-ca') && sourceHasUnreviewedSamples) {
    showToast('Hãy xác nhận hoặc xóa dữ liệu mẫu phụ mổ trước khi chạy tự động.', 'warning', 5000);
    await openClearAssistantsModal();
    return;
  }
  const labels = {
    'tat-ca': 'chạy toàn bộ vào 1 file',
    'bo-sung-phu-mo': 'bổ sung PTV/phụ mổ',
    'chuyen-tieu-phau': 'chuyển thủ thuật sang tiểu phẫu + điền BS',
    'dien-bs': 'điền BS tiểu phẫu'
  };
  setBusy(true, `Đang chạy ${labels[task] || task}...`);
  try {
    await flushAllRows();
    const data = await api('/api/run-task', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file: state.currentFile, task })
    });
    state.currentFile = data.file || data.file_dau_ra || state.currentFile;
    showResult($('taskResult'), data);
    log(`Đã chạy xong: ${labels[task] || task}.`);
    await selectFile(state.currentFile);
    if (state.sheets.includes(state.currentSheet)) $('sheetSelect').value = state.currentSheet;
    await loadSheet();
    await persistCleanState('Đã lưu kết quả tự động vào file làm việc');
  } catch (err) {
    alert(err.message);
    log(`Lỗi chạy tự động: ${err.message}`);
  } finally {
    setBusy(false);
  }
}


async function runNhapPhauThuatSoBo() {
  const soBoInput = $('soBoExcelInput');
  const zipInput = $('soBoZipInput');
  if (!soBoInput.files.length || !zipInput.files.length) {
    alert('Chọn đủ file Excel "Danh sách sơ bộ" và file ZIP ảnh nguồn.');
    return;
  }
  if (!state.currentFile) return;
  setBusy(true, 'Đang nhập dữ liệu từ ảnh sổ phẫu thuật...');
  try {
    await flushAllRows();
    const form = new FormData();
    form.append('file', state.currentFile);
    form.append('soBo', soBoInput.files[0]);
    form.append('zip', zipInput.files[0]);
    const res = await fetch('/api/task-nhap-phau-thuat-so-bo', { method: 'POST', body: form });
    const data = await res.json();
    if (!res.ok || data.ok === false) throw new Error(data.error || 'Nhập dữ liệu lỗi.');
    state.currentFile = data.file || state.currentFile;
    showResult($('taskResult'), data);
    soBoInput.value = '';
    zipInput.value = '';
    if ($('soBoExcelFileName')) $('soBoExcelFileName').textContent = 'Chọn file Excel';
    if ($('soBoZipFileName')) $('soBoZipFileName').textContent = 'Chọn file ZIP';
    await selectFile(state.currentFile);
    if (state.sheets.includes('phauthuat')) $('sheetSelect').value = 'phauthuat';
    await loadSheet('phauthuat');
    await persistCleanState('Đã nhập dữ liệu từ ảnh sổ phẫu thuật');
    const canKiemTra = Number(data.so_dong_can_kiem_tra || 0);
    showToast(
      `Đã thêm ${data.so_dong_them || 0} dòng; ${canKiemTra} dòng cần kiểm tra tên bác sĩ. Nhớ tự điền Loại phẫu thuật và Số tiền cho các dòng mới.`,
      canKiemTra ? 'warning' : 'success',
      8000
    );
    log(`Đã nhập ${data.so_dong_them || 0} dòng từ sổ phẫu thuật (OCR). Báo cáo: ${displayPath(data.report || '')}`);
  } catch (err) {
    alert(err.message);
    log(`Lỗi nhập từ sổ phẫu thuật: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

function soPtFindHeader(candidates) {
  const headers = state.soPhauThuat.headers || [];
  for (const candidate of candidates) {
    const hit = headers.find(h => normalizeKey(h) === normalizeKey(candidate));
    if (hit) return hit;
  }
  return null;
}

function soPtBacSiText(row) {
  const single = soPtFindHeader(['Bác sĩ phẫu thuật', 'Kíp phẫu thuật (đọc theo chuỗi)']);
  if (single && row.values[single]) return String(row.values[single]);
  const parts = [];
  for (let i = 1; i <= 4; i++) {
    const h = soPtFindHeader([`BS phẫu thuật ${i}`]);
    const v = h ? row.values[h] : null;
    if (v && v !== '-') parts.push(v);
  }
  return parts.join(', ');
}

async function loadSoPtRows() {
  if (!state.soPhauThuat.danhSachFile) return;
  const data = await api(`/api/sheet-data?file=${encodeURIComponent(state.soPhauThuat.danhSachFile)}&limit=0`);
  state.soPhauThuat.headers = data.headers || [];
  state.soPhauThuat.rows = data.rows || [];
  state.soPhauThuat.selectedIndex = -1;
  renderSoPtList();
  renderSoPtDetail();
  updateSoPtStats();
}

async function uploadSoPtDanhSach() {
  const input = $('soPtDanhSachInput');
  if (!input.files.length) {
    alert('Chọn file "Danh sách sơ bộ".');
    return;
  }
  setBusy(true, 'Đang tải danh sách sơ bộ...');
  try {
    const form = new FormData();
    form.append('file', input.files[0]);
    const res = await fetch('/api/so-phau-thuat/upload-danh-sach', { method: 'POST', body: form });
    const data = await res.json();
    if (!res.ok || data.ok === false) throw new Error(data.error || 'Tải file lỗi.');
    state.soPhauThuat.danhSachFile = data.file;
    if ($('soPtDanhSachFileName')) $('soPtDanhSachFileName').textContent = data.name;
    await loadSoPtRows();
    showToast(`Đã tải "${data.name}" — ${state.soPhauThuat.rows.length} người bệnh.`, 'success');
    log(`Đã tải danh sách sơ bộ: ${data.name}`);
  } catch (err) {
    alert(err.message);
    log(`Lỗi tải danh sách sơ bộ: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

async function uploadSoPtAnh() {
  const input = $('soPtZipInput');
  if (!input.files.length) {
    alert('Chọn file ZIP ảnh.');
    return;
  }
  setBusy(true, 'Đang giải nén ảnh...');
  try {
    const form = new FormData();
    form.append('zip', input.files[0]);
    const res = await fetch('/api/so-phau-thuat/upload-anh', { method: 'POST', body: form });
    const data = await res.json();
    if (!res.ok || data.ok === false) throw new Error(data.error || 'Tải ảnh lỗi.');
    state.soPhauThuat.imagesLoaded = true;
    if ($('soPtZipFileName')) $('soPtZipFileName').textContent = input.files[0].name;
    if ($('soPtStatus')) $('soPtStatus').textContent = `Đã tải ${data.count} ảnh.`;
    showToast(`Đã tải ${data.count} ảnh vào sổ phẫu thuật.`, 'success');
    log(`Đã tải ${data.count} ảnh vào Sổ phẫu thuật (ảnh).`);
    renderSoPtDetail();
  } catch (err) {
    alert(err.message);
    log(`Lỗi tải ảnh sổ phẫu thuật: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

async function loadDocumentAiConfig() {
  const data = await api('/api/document-ai-config');
  state.documentAiConfig = data.config || {};
  return state.documentAiConfig;
}

function closeDocumentAiConfigModal() {
  $('documentAiConfigModal')?.classList.add('hidden');
  document.body.style.overflow = '';
}

async function openDocumentAiConfigModal() {
  try {
    const config = await loadDocumentAiConfig();
    $('documentAiProjectIdInput').value = config.project_id || '';
    $('documentAiLocationInput').value = config.location || 'us';
    $('documentAiProcessorIdInput').value = config.processor_id || '';
    if ($('documentAiCredentialsStatus')) {
      $('documentAiCredentialsStatus').textContent = config.hasCredentials
        ? 'Đã có file service-account.json.'
        : 'Chưa tải file service-account.json.';
    }
  } catch (err) {
    showToast(err.message, 'error', 5000);
  }
  $('documentAiConfigModal')?.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
  setTimeout(() => $('documentAiProjectIdInput')?.focus(), 60);
}

async function saveDocumentAiConfig() {
  const payload = {
    project_id: $('documentAiProjectIdInput').value.trim(),
    location: $('documentAiLocationInput').value.trim() || 'us',
    processor_id: $('documentAiProcessorIdInput').value.trim(),
  };
  if (!payload.project_id || !payload.processor_id) {
    showToast('Vui lòng nhập Project ID và Processor ID.', 'warning', 4500);
    return;
  }
  setBusy(true, 'Đang lưu cấu hình Document AI...');
  try {
    const credInput = $('documentAiCredentialsInput');
    if (credInput?.files.length) {
      const form = new FormData();
      form.append('file', credInput.files[0]);
      const res = await fetch('/api/document-ai-credentials', { method: 'POST', body: form });
      const credData = await res.json();
      if (!res.ok || credData.ok === false) throw new Error(credData.error || 'Tải file khóa dịch vụ lỗi.');
    }
    const data = await api('/api/document-ai-config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    state.documentAiConfig = data.config || {};
    closeDocumentAiConfigModal();
    showToast('Đã lưu cấu hình Document AI.', 'success');
  } catch (err) {
    showToast(err.message, 'error', 5000);
  } finally {
    setBusy(false);
  }
}

async function runSoPtXuLyAnh() {
  const input = $('soPtAiZipInput');
  if (!input.files.length) {
    alert('Chọn file ZIP ảnh sổ cần xử lý.');
    return;
  }
  try {
    const config = state.documentAiConfig || await loadDocumentAiConfig();
    if (!config.configured) {
      await openDocumentAiConfigModal();
      return;
    }
  } catch (_err) {
    await openDocumentAiConfigModal();
    return;
  }
  const zipFile = input.files[0];
  if ($('soPtAiZipFileName')) $('soPtAiZipFileName').textContent = zipFile.name;
  setBusy(true, 'Đang xử lý ảnh sổ...');
  try {
    if ($('soPtAiStatus')) $('soPtAiStatus').textContent = 'Bước 1/3: Đang giải nén ảnh để xem lại...';
    const formAnh = new FormData();
    formAnh.append('zip', zipFile);
    const resAnh = await fetch('/api/so-phau-thuat/upload-anh', { method: 'POST', body: formAnh });
    const dataAnh = await resAnh.json();
    if (!resAnh.ok || dataAnh.ok === false) throw new Error(dataAnh.error || 'Giải nén ảnh lỗi.');
    state.soPhauThuat.imagesLoaded = true;

    if ($('soPtAiStatus')) $('soPtAiStatus').textContent = 'Bước 2/3: Đang chạy OCR (Google Document AI) — có thể mất vài phút...';
    const formOcr = new FormData();
    formOcr.append('zip', zipFile);
    const resOcr = await fetch('/api/so-phau-thuat/chay-ocr', { method: 'POST', body: formOcr });
    const dataOcr = await resOcr.json();
    if (!resOcr.ok || dataOcr.ok === false) throw new Error(dataOcr.error || 'Chạy OCR lỗi.');

    if ($('soPtAiStatus')) $('soPtAiStatus').textContent = 'Bước 3/3: Đang tách cột thành dữ liệu Excel...';
    const dataTach = await api('/api/so-phau-thuat/tach-cot', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({}),
    });
    state.soPhauThuat.danhSachFile = dataTach.file;
    if ($('soPtDanhSachFileName')) $('soPtDanhSachFileName').textContent = `Tự tách từ OCR (${dataTach.count || 0} dòng — CẦN KIỂM TRA)`;
    await loadSoPtRows();

    if ($('soPtAiStatus')) {
      $('soPtAiStatus').textContent = `Xong: OCR ${dataAnh.count || 0} ảnh, tách được ${dataTach.count || 0} dòng — TẤT CẢ đều cần xem ảnh gốc để xác minh trước khi dùng.`;
    }
    showToast(`Đã xử lý xong ${dataAnh.count || 0} ảnh — ${dataTach.count || 0} dòng cần kiểm tra lại.`, 'warning', 6000);
    log(`Đã xử lý ảnh sổ tự động: ${dataAnh.count || 0} ảnh, ${dataTach.count || 0} dòng.`);
  } catch (err) {
    alert(err.message);
    if ($('soPtAiStatus')) $('soPtAiStatus').textContent = '';
    log(`Lỗi xử lý ảnh sổ tự động: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

function renderSoPtList() {
  const tbody = $('soPtTableBody');
  if (!tbody) return;
  const query = normalizeKey($('soPtSearchInput')?.value || '');
  const hNgay = soPtFindHeader(['Ngày ghi sổ', 'Ngày']);
  const hHoTen = soPtFindHeader(['Họ tên người bệnh', 'Họ và tên']);
  const hChanDoan = soPtFindHeader(['Chẩn đoán đọc từ sổ', 'Chẩn đoán']);
  const hPhuongPhap = soPtFindHeader(['Phương pháp phẫu thuật', 'Phương pháp']);
  const hAnh = soPtFindHeader(['Ảnh nguồn', 'Ảnh nguồn trong ZIP']);

  const filtered = (state.soPhauThuat.rows || []).filter(row => {
    if (!query) return true;
    const haystack = normalizeKey(Object.values(row.values || {}).join(' '));
    return haystack.includes(query);
  });
  state.soPhauThuat.filtered = filtered;
  if (state.soPhauThuat.selectedIndex >= filtered.length) state.soPhauThuat.selectedIndex = -1;

  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="5" class="so-pt-detail-empty">${state.soPhauThuat.rows.length ? 'Không tìm thấy dòng phù hợp.' : 'Chưa tải file "Danh sách sơ bộ".'}</td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map((row, idx) => {
    const activeClass = idx === state.soPhauThuat.selectedIndex ? ' class="active"' : '';
    const ngay = hNgay ? (row.values[hNgay] ?? '') : '';
    const hoTen = hHoTen ? (row.values[hHoTen] ?? '') : '';
    const chanDoanPp = [hChanDoan ? row.values[hChanDoan] : '', hPhuongPhap ? row.values[hPhuongPhap] : '']
      .filter(v => v !== null && v !== undefined && v !== '')
      .join(' — ');
    const bacSi = soPtBacSiText(row);
    const anh = hAnh ? (row.values[hAnh] ?? '') : '';
    return `<tr${activeClass} data-idx="${idx}"><td>${escapeHtml(ngay)}</td><td>${escapeHtml(hoTen)}</td><td>${escapeHtml(chanDoanPp)}</td><td>${escapeHtml(bacSi)}</td><td>${escapeHtml(anh)}</td></tr>`;
  }).join('');

  tbody.querySelectorAll('tr[data-idx]').forEach(tr => {
    tr.addEventListener('click', () => selectSoPtRow(Number(tr.dataset.idx)));
  });
}

function selectSoPtRow(idx) {
  state.soPhauThuat.selectedIndex = idx;
  renderSoPtList();
  renderSoPtDetail();
}

function renderSoPtDetail() {
  const box = $('soPtDetail');
  if (!box) return;
  const row = state.soPhauThuat.filtered?.[state.soPhauThuat.selectedIndex];
  if (!row) {
    box.innerHTML = '<div class="so-pt-detail-empty">Bấm vào một người bệnh trong danh sách để xem ảnh gốc trang sổ.</div>';
    return;
  }

  const hAnh = soPtFindHeader(['Ảnh nguồn', 'Ảnh nguồn trong ZIP']);
  const anhName = hAnh ? String(row.values[hAnh] || '').trim() : '';

  box.innerHTML = '';

  const imageWrap = document.createElement('div');
  if (anhName) {
    imageWrap.className = 'so-pt-image-wrap';
    const img = document.createElement('img');
    img.src = `/api/so-phau-thuat/anh/${encodeURIComponent(anhName)}`;
    img.alt = anhName;
    img.addEventListener('error', () => {
      imageWrap.className = 'so-pt-image-missing';
      imageWrap.textContent = `Chưa tải được ảnh "${anhName}" — bấm "Tải ảnh" ở trên rồi thử lại.`;
    });
    imageWrap.appendChild(img);
  } else {
    imageWrap.className = 'so-pt-image-missing';
    imageWrap.textContent = 'Dòng này chưa có tên ảnh nguồn.';
  }
  box.appendChild(imageWrap);

  const skipHeaders = new Set([hAnh].filter(Boolean));
  (state.soPhauThuat.headers || []).forEach(h => {
    if (skipHeaders.has(h)) return;
    const value = row.values[h];
    if (value === null || value === undefined || value === '') return;
    const field = document.createElement('div');
    field.className = 'so-pt-field';
    const label = document.createElement('strong');
    label.textContent = h;
    field.appendChild(label);
    field.appendChild(document.createTextNode(String(value)));
    box.appendChild(field);
  });
}

function updateSoPtStats() {
  const box = $('soPtStats');
  if (!box) return;
  const total = state.soPhauThuat.rows.length;
  box.innerHTML = `<div class="overview-stat"><strong>${total}</strong><span>người bệnh</span></div>`;
}

async function runSoPtDoiChieu() {
  if (!state.currentFile) {
    alert('Chưa có file đang mở ở tab "Nhập dữ liệu".');
    return;
  }
  if (!state.soPhauThuat.danhSachFile) {
    alert('Chưa tải file "Danh sách sơ bộ" ở trên.');
    return;
  }
  setBusy(true, 'Đang đối chiếu với sheet phauthuat...');
  if ($('soPtDoiChieuStatus')) $('soPtDoiChieuStatus').textContent = 'Đang đối chiếu...';
  try {
    const data = await api(
      `/api/so-phau-thuat/doi-chieu?file=${encodeURIComponent(state.currentFile)}&soBo=${encodeURIComponent(state.soPhauThuat.danhSachFile)}`
    );
    state.soPtDoiChieu = data.goiY || [];
    state.soPtDoiChieuOpenRows = new Set();
    renderSoPtDoiChieuList();
    const count = state.soPtDoiChieu.length;
    if ($('soPtDoiChieuStatus')) {
      $('soPtDoiChieuStatus').textContent = count
        ? `Tìm được ${count} dòng thiếu PTV chính có gợi ý từ ảnh.`
        : 'Không có dòng nào thiếu PTV chính khớp được với dữ liệu ảnh.';
    }
    log(`Đối chiếu phauthuat với ảnh: ${count} gợi ý.`);
  } catch (err) {
    alert(err.message);
    log(`Lỗi đối chiếu phauthuat với ảnh: ${err.message}`);
    if ($('soPtDoiChieuStatus')) $('soPtDoiChieuStatus').textContent = '';
  } finally {
    setBusy(false);
  }
}

function renderSoPtDoiChieuList() {
  const box = $('soPtDoiChieuList');
  if (!box) return;
  if (!state.soPtDoiChieu.length) {
    box.innerHTML = '';
    return;
  }
  box.innerHTML = state.soPtDoiChieu.map(item => {
    const phuMoText = (item.goiYPhuMo || []).filter(Boolean).join(', ') || '(không có)';
    const warn = item.soUngVienCungTen > 1
      ? `<div class="so-pt-warn">Có ${item.soUngVienCungTen} dòng ảnh cùng ngày + họ tên, độ khớp phương pháp ${item.doKhopPhuongPhap ?? '?'} — kiểm tra kỹ trước khi áp dụng.</div>`
      : '';
    const isOpen = state.soPtDoiChieuOpenRows.has(item.row);
    return `
      <div class="so-pt-goi-y-card" data-row="${item.row}">
        <div class="so-pt-goi-y-head-row">
          <strong>${escapeHtml(item.hoTen)}</strong>
          <span>${escapeHtml(item.ngay || '')} · Dòng phauthuat ${item.row}</span>
        </div>
        <div class="so-pt-goi-y-body">
          <div>Chẩn đoán/PP đang có: ${escapeHtml(item.chanDoanPhuongPhap || '(trống)')}</div>
          <div>Đọc từ ảnh: ${escapeHtml(item.soBoChanDoan || '')}${item.soBoPhuongPhap ? ' — ' + escapeHtml(item.soBoPhuongPhap) : ''}</div>
          <div>Gợi ý PTV chính: <strong>${escapeHtml(item.goiYPtvChinh || '(không đọc được)')}</strong></div>
          <div>Gợi ý phụ mổ: ${escapeHtml(phuMoText)}</div>
          ${warn}
        </div>
        <div class="so-pt-goi-y-actions">
          <button class="btn subtle" type="button" data-action="xem-anh" data-row="${item.row}">${isOpen ? 'Ẩn ảnh' : 'Xem ảnh'}</button>
          <button class="btn primary" type="button" data-action="ap-dung" data-row="${item.row}">Áp dụng vào dòng ${item.row}</button>
        </div>
        <div class="so-pt-goi-y-image${isOpen ? '' : ' hidden'}" data-image-for="${item.row}"></div>
      </div>
    `;
  }).join('');

  state.soPtDoiChieuOpenRows.forEach(row => fillSoPtGoiYImage(row));
}

function fillSoPtGoiYImage(row) {
  const item = state.soPtDoiChieu.find(x => x.row === row);
  const slot = document.querySelector(`.so-pt-goi-y-image[data-image-for="${row}"]`);
  if (!item || !slot) return;
  if (!item.anhNguon) {
    slot.innerHTML = '<div class="so-pt-image-missing">Dòng ảnh này không có tên ảnh nguồn.</div>';
    return;
  }
  slot.innerHTML = '';
  const img = document.createElement('img');
  img.src = `/api/so-phau-thuat/anh/${encodeURIComponent(item.anhNguon)}`;
  img.alt = item.anhNguon;
  img.addEventListener('error', () => {
    slot.innerHTML = `<div class="so-pt-image-missing">Chưa có ảnh "${escapeHtml(item.anhNguon)}" — bấm "Tải ảnh" ở trên trước.</div>`;
  });
  slot.appendChild(img);
}

function toggleSoPtGoiYAnh(row) {
  if (state.soPtDoiChieuOpenRows.has(row)) state.soPtDoiChieuOpenRows.delete(row);
  else state.soPtDoiChieuOpenRows.add(row);
  renderSoPtDoiChieuList();
}

async function apDungSoPtGoiY(row) {
  const item = state.soPtDoiChieu.find(x => x.row === row);
  if (!item) return;
  if (!state.currentFile) {
    alert('Chưa có file đang mở ở tab "Nhập dữ liệu".');
    return;
  }

  const data = {};
  if (item.goiYPtvChinh) data['PTV chính'] = item.goiYPtvChinh;
  (item.goiYPhuMo || []).forEach((value, i) => {
    if (value) data[`Phụ mổ ${i + 1}`] = value;
  });
  if (!Object.keys(data).length) {
    alert('Không có gợi ý nào để áp dụng cho dòng này.');
    return;
  }
  if (!confirm(`Áp dụng vào dòng ${item.row} (${item.hoTen})?\nPTV chính: ${data['PTV chính'] || '(giữ trống)'}\nPhụ mổ: ${(item.goiYPhuMo || []).filter(Boolean).join(', ') || '(không có)'}`)) {
    return;
  }

  setBusy(true, `Đang áp dụng vào dòng ${item.row}...`);
  try {
    const res = await api('/api/update-rows', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file: state.currentFile,
        sheet: 'phauthuat',
        rows: [{ rowNumber: item.row, data }],
        overwrite: false,
        suffix: '_NHAP_LIEU',
      }),
    });
    state.currentFile = res.file || state.currentFile;
    addFileOptionIfMissing(state.currentFile);
    state.soPtDoiChieu = state.soPtDoiChieu.filter(x => x.row !== row);
    state.soPtDoiChieuOpenRows.delete(row);
    renderSoPtDoiChieuList();
    showToast(`Đã điền dòng ${item.row} (${item.hoTen}).`, 'success');
    log(`Đã áp dụng gợi ý từ ảnh cho dòng ${item.row} sheet phauthuat (từ ${res.file}).`);
    if (state.currentSheet === 'phauthuat' && document.getElementById('dataView') && !document.getElementById('dataView').classList.contains('hidden')) {
      await loadSheet();
    }
  } catch (err) {
    alert(err.message);
    log(`Lỗi áp dụng gợi ý dòng ${item.row}: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

function switchMainTab(tabName) {
  const isData = tabName === 'data';
  const isStaff = tabName === 'staff';
  const isCls = tabName === 'cls';
  const isSoPt = tabName === 'sopt';
  $('dataView')?.classList.toggle('hidden', !isData);
  $('staffView')?.classList.toggle('hidden', !isStaff);
  $('clsView')?.classList.toggle('hidden', !isCls);
  $('soPhauThuatView')?.classList.toggle('hidden', !isSoPt);
  $('dataTabBtn')?.classList.toggle('active', isData);
  $('staffTabBtn')?.classList.toggle('active', isStaff);
  $('clsTabBtn')?.classList.toggle('active', isCls);
  $('soPhauThuatTabBtn')?.classList.toggle('active', isSoPt);
  hideSuggestPanel();
  if (isStaff) {
    renderStaffManager();
    resetStaffForm();
    document.title = 'Quản lý nhân sự | PM CTCH';
  } else if (isCls) {
    renderClsManager();
    resetClsForm();
    document.title = 'Quản lý tên CLS | PM CTCH';
  } else if (isSoPt) {
    renderSoPtList();
    renderSoPtDetail();
    updateSoPtStats();
    document.title = 'Sổ phẫu thuật (ảnh) | PM CTCH';
  } else {
    updateWorkspaceIdentity();
  }
}

function staffRoleLabel(role) {
  return {
    bac_si: 'Bác sĩ',
    bsnt: 'Bác sĩ nội trú',
    dieu_duong: 'Điều dưỡng',
    ktv: 'Kỹ thuật viên',
  }[role] || role || 'Khác';
}

function renderStaffStats() {
  const host = $('staffStats');
  if (!host) return;
  const active = state.staff.filter(item => item.active !== false);
  const counts = {
    total: state.staff.length,
    bac_si: active.filter(item => item.vaiTro === 'bac_si').length,
    bsnt: active.filter(item => item.vaiTro === 'bsnt').length,
    dieu_duong: active.filter(item => item.vaiTro === 'dieu_duong').length,
    ktv: active.filter(item => item.vaiTro === 'ktv').length,
  };
  host.innerHTML = [
    ['Tổng nhân sự', counts.total],
    ['Bác sĩ', counts.bac_si],
    ['BSNT', counts.bsnt],
    ['Điều dưỡng', counts.dieu_duong],
    ['Kỹ thuật viên', counts.ktv],
  ].map(([label, value]) => `<div class="staff-stat"><strong>${value}</strong><span>${label}</span></div>`).join('');
}

function renderStaffManager() {
  renderStaffStats();
  const body = $('staffTableBody');
  if (!body) return;
  const query = normalizeKey($('staffSearchInput')?.value || '');
  const role = $('staffRoleFilter')?.value || 'all';
  const rows = state.staff
    .filter(item => role === 'all' || item.vaiTro === role)
    .filter(item => !query || normalizeKey(`${item.biDanh} ${item.hoTen} ${item.nhom}`).includes(query))
    .sort((a, b) => {
      const activeDiff = Number(b.active !== false) - Number(a.active !== false);
      if (activeDiff) return activeDiff;
      const roleDiff = staffRoleLabel(a.vaiTro).localeCompare(staffRoleLabel(b.vaiTro), 'vi');
      if (roleDiff) return roleDiff;
      return String(a.biDanh).localeCompare(String(b.biDanh), 'vi');
    });

  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="6" class="empty-state">Không có nhân sự phù hợp bộ lọc.</td></tr>';
    return;
  }

  body.innerHTML = rows.map(item => `
    <tr>
      <td class="alias-cell">${escapeHtml(item.biDanh)}</td>
      <td>${escapeHtml(item.hoTen)}</td>
      <td><span class="role-badge ${escapeHtml(item.vaiTro)}">${escapeHtml(staffRoleLabel(item.vaiTro))}</span></td>
      <td>${escapeHtml(item.nhom || '')}</td>
      <td><span class="active-badge ${item.active !== false ? 'active' : 'inactive'}">${item.active !== false ? 'Đang dùng' : 'Tạm ngưng'}</span></td>
      <td><button class="edit-staff-btn" type="button" data-staff-id="${escapeHtml(item.id)}">Sửa</button></td>
    </tr>
  `).join('');

  body.querySelectorAll('[data-staff-id]').forEach(btn => {
    btn.addEventListener('click', () => editStaff(btn.dataset.staffId));
  });
}

function defaultGroupForRole(role) {
  return {
    bac_si: 'Bác sĩ',
    bsnt: 'BSNT',
    dieu_duong: 'Điều dưỡng',
    ktv: 'KTV',
  }[role] || '';
}

function resetStaffForm() {
  if (!$('staffIdInput')) return;
  $('staffIdInput').value = '';
  $('staffNameInput').value = '';
  $('staffAliasInput').value = '';
  $('staffRoleInput').value = 'bac_si';
  $('staffGroupInput').value = 'Bác sĩ';
  $('staffActiveInput').checked = true;
  $('staffFormTitle').textContent = 'Thêm nhân sự';
  $('deleteStaffBtn').classList.add('hidden');
  $('staffNameInput').focus();
}

function editStaff(id) {
  const item = state.staff.find(row => row.id === id);
  if (!item) return;
  $('staffIdInput').value = item.id;
  $('staffNameInput').value = item.hoTen || '';
  $('staffAliasInput').value = item.biDanh || '';
  $('staffRoleInput').value = item.vaiTro || 'bac_si';
  $('staffGroupInput').value = item.nhom || defaultGroupForRole(item.vaiTro);
  $('staffActiveInput').checked = item.active !== false;
  $('staffFormTitle').textContent = `Chỉnh sửa ${item.biDanh}`;
  $('deleteStaffBtn').classList.remove('hidden');
  $('staffNameInput').focus();
}

async function persistStaffList(nextStaff, successMessage) {
  const data = await api('/api/staff-list', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ staff: nextStaff }),
  });
  state.staff = data.staff || nextStaff;
  renderStaffManager();
  const hasAutoAdjust = Array.isArray(data.autoAdjusted) && data.autoAdjusted.length > 0;
  const message = hasAutoAdjust ? data.message : (successMessage || 'Đã lưu danh sách nhân sự.');
  showToast(message, hasAutoAdjust ? 'warning' : 'success', hasAutoAdjust ? 7000 : 3600);
  log(message);
}

async function saveStaffFromForm() {
  const name = $('staffNameInput').value.trim();
  const alias = $('staffAliasInput').value.trim().toUpperCase();
  const role = $('staffRoleInput').value;
  const group = $('staffGroupInput').value.trim() || defaultGroupForRole(role);
  const id = $('staffIdInput').value || (window.crypto?.randomUUID ? crypto.randomUUID() : `staff_${Date.now()}`);
  if (!name || !alias) {
    showToast('Cần nhập đủ họ tên và bí danh.', 'warning');
    return;
  }
  const item = { id, hoTen: name, biDanh: alias, vaiTro: role, nhom: group, active: $('staffActiveInput').checked };
  const existingIndex = state.staff.findIndex(row => row.id === id);
  const next = state.staff.map(row => ({ ...row }));
  if (existingIndex >= 0) next[existingIndex] = item;
  else next.push(item);

  setBusy(true, 'Đang lưu nhân sự...');
  try {
    await persistStaffList(next, existingIndex >= 0 ? `Đã cập nhật ${alias}.` : `Đã thêm ${alias}.`);
    resetStaffForm();
  } catch (err) {
    showToast(err.message, 'error', 5000);
  } finally {
    setBusy(false);
  }
}

async function deleteCurrentStaff() {
  const id = $('staffIdInput').value;
  const item = state.staff.find(row => row.id === id);
  if (!item) return;
  if (!confirm(`Xóa ${item.biDanh} — ${item.hoTen} khỏi danh sách nhân sự?`)) return;
  setBusy(true, 'Đang xóa nhân sự...');
  try {
    await persistStaffList(state.staff.filter(row => row.id !== id), `Đã xóa ${item.biDanh}.`);
    resetStaffForm();
  } catch (err) {
    showToast(err.message, 'error', 5000);
  } finally {
    setBusy(false);
  }
}



async function persistClsList(nextList, successMessage) {
  const data = await api('/api/cls-list', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ cls: nextList }),
  });
  state.clsList = data.cls || nextList;
  renderClsManager();
  if (successMessage) showToast(successMessage, 'success');
}

function renderClsStats() {
  const host = $('clsStats');
  if (!host) return;
  const activeCount = state.clsList.filter(item => item.active !== false).length;
  host.innerHTML = [
    ['Tổng tên CLS', state.clsList.length],
    ['Đang chuyển', activeCount],
    ['Tạm ngưng', Math.max(0, state.clsList.length - activeCount)],
  ].map(([label, value]) => `<div class="staff-stat"><strong>${value}</strong><span>${label}</span></div>`).join('');
}

function renderClsManager() {
  renderClsStats();
  const body = $('clsTableBody');
  if (!body) return;
  const query = normalizeKey($('clsSearchInput')?.value || '');
  const status = $('clsStatusFilter')?.value || 'all';
  const rows = state.clsList
    .filter(item => status === 'all' || (status === 'active' ? item.active !== false : item.active === false))
    .filter(item => !query || normalizeKey(item.tenCls || '').includes(query))
    .sort((a, b) => {
      const activeDiff = Number(b.active !== false) - Number(a.active !== false);
      if (activeDiff) return activeDiff;
      return String(a.tenCls || '').localeCompare(String(b.tenCls || ''), 'vi');
    });

  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="3" class="empty-state">Không có tên CLS phù hợp bộ lọc.</td></tr>';
    return;
  }
  body.innerHTML = rows.map(item => `
    <tr>
      <td class="cls-name-cell">${escapeHtml(item.tenCls || '')}</td>
      <td><span class="active-badge ${item.active !== false ? 'active' : 'inactive'}">${item.active !== false ? 'Đang chuyển' : 'Tạm ngưng'}</span></td>
      <td><button class="edit-staff-btn" type="button" data-cls-id="${escapeHtml(item.id)}">Sửa</button></td>
    </tr>
  `).join('');
  body.querySelectorAll('[data-cls-id]').forEach(btn => {
    btn.addEventListener('click', () => editCls(btn.dataset.clsId));
  });
}

function resetClsForm() {
  if (!$('clsIdInput')) return;
  $('clsIdInput').value = '';
  $('clsNameInput').value = '';
  $('clsActiveInput').checked = true;
  $('clsFormTitle').textContent = 'Thêm tên CLS';
  $('deleteClsBtn').classList.add('hidden');
  $('clsNameInput').focus();
}

function editCls(id) {
  const item = state.clsList.find(row => row.id === id);
  if (!item) return;
  $('clsIdInput').value = item.id;
  $('clsNameInput').value = item.tenCls || '';
  $('clsActiveInput').checked = item.active !== false;
  $('clsFormTitle').textContent = 'Chỉnh sửa tên CLS';
  $('deleteClsBtn').classList.remove('hidden');
  $('clsNameInput').focus();
}

async function saveClsFromForm() {
  const name = $('clsNameInput').value.trim();
  if (!name) {
    showToast('Cần nhập tên CLS.', 'warning');
    return;
  }
  const id = $('clsIdInput').value || (window.crypto?.randomUUID ? crypto.randomUUID() : `cls_${Date.now()}`);
  const item = { id, tenCls: name, active: $('clsActiveInput').checked };
  const existingIndex = state.clsList.findIndex(row => row.id === id);
  const next = state.clsList.map(row => ({ ...row }));
  if (existingIndex >= 0) next[existingIndex] = item;
  else next.push(item);

  setBusy(true, 'Đang lưu tên CLS...');
  try {
    await persistClsList(next, existingIndex >= 0 ? 'Đã cập nhật tên CLS.' : 'Đã thêm tên CLS.');
    resetClsForm();
  } catch (err) {
    showToast(err.message, 'error', 5000);
  } finally {
    setBusy(false);
  }
}

async function deleteCurrentCls() {
  const id = $('clsIdInput').value;
  const item = state.clsList.find(row => row.id === id);
  if (!item) return;
  if (!confirm(`Xóa tên CLS “${item.tenCls}” khỏi danh mục chuyển tiểu phẫu?`)) return;
  setBusy(true, 'Đang xóa tên CLS...');
  try {
    await persistClsList(state.clsList.filter(row => row.id !== id), 'Đã xóa tên CLS.');
    resetClsForm();
  } catch (err) {
    showToast(err.message, 'error', 5000);
  } finally {
    setBusy(false);
  }
}

async function finalizeAndDownload(event) {
  event?.preventDefault();
  if (!state.currentFile) {
    showToast('Chưa chọn file Excel.', 'warning');
    return;
  }
  setBusy(true, 'Đang hoàn tất công thức và dòng Tổng Cộng...');
  try {
    await flushAllRows();
    const data = await api('/api/finalize-workbook', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file: state.currentFile }),
    });
    state.currentFile = data.file || state.currentFile;
    await persistCleanState('Đã hoàn tất dữ liệu trước khi tải Excel');
    const link = document.createElement('a');
    link.href = `/download?file=${encodeURIComponent(state.currentFile)}`;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
    const filled = (data.sheets || []).reduce((sum, item) => sum + Number(item.formulasWritten || 0), 0);
    showToast(`Đã hoàn tất file và tạo ${filled} công thức tiền.`, 'success', 4800);
  } catch (err) {
    showToast(`Không thể hoàn tất file: ${err.message}`, 'error', 6500);
    log(`Lỗi hoàn tất trước khi tải: ${err.message}`);
  } finally {
    setBusy(false);
  }
}

function bindEvents() {
  $('dataTabBtn')?.addEventListener('click', () => switchMainTab('data'));
  $('staffTabBtn')?.addEventListener('click', () => switchMainTab('staff'));
  $('clsTabBtn')?.addEventListener('click', () => switchMainTab('cls'));
  $('soPhauThuatTabBtn')?.addEventListener('click', () => switchMainTab('sopt'));
  $('soPtDanhSachBtn')?.addEventListener('click', uploadSoPtDanhSach);
  $('soPtZipBtn')?.addEventListener('click', uploadSoPtAnh);
  $('soPtDanhSachInput')?.addEventListener('change', e => {
    const file = e.target.files?.[0];
    if ($('soPtDanhSachFileName')) $('soPtDanhSachFileName').textContent = file ? file.name : 'Chọn file "Danh sách sơ bộ"';
  });
  $('soPtZipInput')?.addEventListener('change', e => {
    const file = e.target.files?.[0];
    if ($('soPtZipFileName')) $('soPtZipFileName').textContent = file ? file.name : 'Chọn file ZIP ảnh';
  });
  $('soPtSearchInput')?.addEventListener('input', () => renderSoPtList());
  $('soPtDoiChieuBtn')?.addEventListener('click', runSoPtDoiChieu);
  $('soPtAiConfigBtn')?.addEventListener('click', e => { e.preventDefault(); openDocumentAiConfigModal(); });
  $('soPtAiXuLyBtn')?.addEventListener('click', runSoPtXuLyAnh);
  $('soPtAiZipInput')?.addEventListener('change', e => {
    const file = e.target.files?.[0];
    if ($('soPtAiZipFileName')) $('soPtAiZipFileName').textContent = file ? file.name : 'Chọn ZIP ảnh sổ';
  });
  $('closeDocumentAiConfigBtn')?.addEventListener('click', closeDocumentAiConfigModal);
  $('cancelDocumentAiConfigBtn')?.addEventListener('click', closeDocumentAiConfigModal);
  $('saveDocumentAiConfigBtn')?.addEventListener('click', saveDocumentAiConfig);
  $('documentAiCredentialsInput')?.addEventListener('change', e => {
    const file = e.target.files?.[0];
    if ($('documentAiCredentialsFileName')) $('documentAiCredentialsFileName').textContent = file ? file.name : 'Chọn file khóa dịch vụ (service-account.json)';
  });
  $('documentAiConfigModal')?.addEventListener('mousedown', e => {
    if (e.target === $('documentAiConfigModal')) closeDocumentAiConfigModal();
  });
  $('soPtDoiChieuList')?.addEventListener('click', e => {
    const btn = e.target.closest('button[data-action]');
    if (!btn) return;
    const row = Number(btn.dataset.row);
    if (btn.dataset.action === 'xem-anh') toggleSoPtGoiYAnh(row);
    else if (btn.dataset.action === 'ap-dung') apDungSoPtGoiY(row);
  });
  $('newStaffBtn')?.addEventListener('click', resetStaffForm);
  $('saveStaffBtn')?.addEventListener('click', saveStaffFromForm);
  $('cancelStaffEditBtn')?.addEventListener('click', resetStaffForm);
  $('deleteStaffBtn')?.addEventListener('click', deleteCurrentStaff);
  $('staffSearchInput')?.addEventListener('input', renderStaffManager);
  $('staffRoleFilter')?.addEventListener('change', renderStaffManager);
  $('staffRoleInput')?.addEventListener('change', event => {
    if (!$('staffIdInput').value || !$('staffGroupInput').value.trim()) {
      $('staffGroupInput').value = defaultGroupForRole(event.target.value);
    }
  });

  $('newClsBtn')?.addEventListener('click', resetClsForm);
  $('saveClsBtn')?.addEventListener('click', saveClsFromForm);
  $('cancelClsEditBtn')?.addEventListener('click', resetClsForm);
  $('deleteClsBtn')?.addEventListener('click', deleteCurrentCls);
  $('clsSearchInput')?.addEventListener('input', renderClsManager);
  $('clsStatusFilter')?.addEventListener('change', renderClsManager);
  $('downloadCurrent')?.addEventListener('click', finalizeAndDownload);

  $('refreshFilesBtn').addEventListener('click', () => {
    loadFiles().catch(err => {
      alert(err.message);
      log(`Lỗi tải danh sách file: ${err.message}`);
    });
  });
  $('fileSelect').addEventListener('change', async e => {
    const requestedFile = e.target.value;
    setBusy(true, 'Đang đổi file...');
    try {
      await selectFile(requestedFile);
      await loadSheet();
    } catch (err) {
      if (state.currentFile) $('fileSelect').value = state.currentFile;
      alert(err.message);
      log(`Không đổi được file: ${err.message}`);
    } finally {
      setBusy(false);
    }
  });
  $('uploadBtn').addEventListener('click', uploadFile);
  $('theoSoUploadBtn')?.addEventListener('click', uploadTheoSoFiles);
  $('theoSoUploadInput')?.addEventListener('change', e => {
    const files = Array.from(e.target.files || []);
    if ($('theoSoUploadFileName')) {
      $('theoSoUploadFileName').textContent = files.length
        ? (files.length === 1 ? files[0].name : `${files.length} file đã chọn`)
        : 'Chọn file Excel';
    }
  });
  $('nhapPhauThuatSoBoBtn')?.addEventListener('click', runNhapPhauThuatSoBo);
  $('soBoExcelInput')?.addEventListener('change', e => {
    const file = e.target.files?.[0];
    if ($('soBoExcelFileName')) $('soBoExcelFileName').textContent = file ? file.name : 'Chọn file Excel';
  });
  $('soBoZipInput')?.addEventListener('change', e => {
    const file = e.target.files?.[0];
    if ($('soBoZipFileName')) $('soBoZipFileName').textContent = file ? file.name : 'Chọn file ZIP';
  });
  $('loadSheetBtn').addEventListener('click', () => {
    loadSheet().catch(err => alert(err.message));
  });
  $('sheetSelect').addEventListener('change', async e => {
    const requestedSheet = e.target.value;
    try {
      await loadSheet(requestedSheet);
    } catch (err) {
      if (state.currentSheet) $('sheetSelect').value = state.currentSheet;
      alert(err.message);
    }
  });
  $('viewModeSelect').addEventListener('change', () => {
    state.visibleHeaders = chooseDisplayHeaders();
    renderDatalists();
    renderTable();
    saveDraftLocal(currentDraftPayload());
    scheduleDraftSave();
  });
  $('searchBtn').addEventListener('click', () => loadSheet().catch(err => alert(err.message)));
  $('clearSearchBtn').addEventListener('click', () => {
    $('searchInput').value = '';
    $('missingOnly').checked = false;
    loadSheet().catch(err => alert(err.message));
  });
  $('searchInput').addEventListener('keydown', e => {
    if (e.key === 'Enter') loadSheet().catch(err => alert(err.message));
  });
  $('missingOnly').addEventListener('change', () => loadSheet().catch(err => alert(err.message)));
  $('addNewRowBtn').addEventListener('click', e => { e.preventDefault(); startNewRow(); });
  $('commitNewRowBtn').addEventListener('click', e => { e.preventDefault(); insertNewRow(); });
  $('copySelectedBtn').addEventListener('click', e => { e.preventDefault(); copySelectedRow(); });
  $('saveAllBtn')?.addEventListener('click', e => { e.preventDefault(); saveAllRows(); });
  $('deleteSelectedBtn').addEventListener('click', e => { e.preventDefault(); deleteSelectedRow(); });
  $('restoreDraftBtn').addEventListener('click', e => { e.preventDefault(); restoreBestDraftIfAny(false); });
  $('clearDraftBtn').addEventListener('click', e => { e.preventDefault(); clearAutosaveDraft(); });
  $('refreshClsAndEmrBtn')?.addEventListener('click', e => { e.preventDefault(); updateMissingClsAndDoctors(); });
  $('emrFillAllBtn')?.addEventListener('click', e => { e.preventDefault(); runEmrFill(false); });
  $('emrFillSelectedBtn')?.addEventListener('click', e => { e.preventDefault(); runEmrFill(true); });
  $('openEmrConfigBtn')?.addEventListener('click', e => { e.preventDefault(); openEmrConfigModal(); });
  $('closeEmrConfigBtn')?.addEventListener('click', closeEmrConfigModal);
  $('cancelEmrConfigBtn')?.addEventListener('click', closeEmrConfigModal);
  $('saveEmrConfigBtn')?.addEventListener('click', saveEmrConfig);
  $('clearAssistantsBtn').addEventListener('click', e => { e.preventDefault(); openClearAssistantsModal(); });
  $('openClearAssistantsBtn').addEventListener('click', e => { e.preventDefault(); openClearAssistantsModal(); });
  $('dismissSampleNoticeBtn').addEventListener('click', e => {
    e.preventDefault();
    state.dismissedSampleNoticeFor = `${state.currentFile}|${state.currentSheet}`;
    updateSampleDataNotice();
  });
  $('closeClearModalBtn').addEventListener('click', closeClearAssistantsModal);
  $('cancelClearAssistantsBtn').addEventListener('click', closeClearAssistantsModal);
  $('confirmClearAssistantsBtn').addEventListener('click', confirmClearAssistants);
  $('clearConfirmCheck').addEventListener('change', updateClearPreviewSummary);
  document.querySelectorAll('.clear-column-check').forEach(input => input.addEventListener('change', updateClearPreviewSummary));
  $('uploadInput').addEventListener('change', e => {
    const file = e.target.files?.[0];
    $('uploadFileName').textContent = file ? file.name : 'Chọn file Excel';
  });
  $('clearAssistantsModal').addEventListener('mousedown', e => {
    if (e.target === $('clearAssistantsModal')) closeClearAssistantsModal();
  });
  $('emrConfigModal')?.addEventListener('mousedown', e => {
    if (e.target === $('emrConfigModal')) closeEmrConfigModal();
  });
  window.addEventListener('beforeunload', () => {
    saveDraftLocal(currentDraftPayload());
  });
  document.addEventListener('mousedown', event => {
    const panel = $('suggestPanel');
    if (!panel || panel.classList.contains('hidden')) return;
    if (panel.contains(event.target) || event.target === state.focusedInput) return;
    hideSuggestPanel();
  });
  document.addEventListener('keydown', event => {
    if (event.key !== 'Escape') return;
    if (!$('clearAssistantsModal').classList.contains('hidden')) closeClearAssistantsModal();
    if (!$('emrConfigModal')?.classList.contains('hidden')) closeEmrConfigModal();
    if (!$('documentAiConfigModal')?.classList.contains('hidden')) closeDocumentAiConfigModal();
  });
  document.querySelectorAll('[data-task]').forEach(btn => {
    btn.addEventListener('click', () => runTask(btn.dataset.task));
  });
}

async function loadHeaderAliases() {
  try {
    const data = await api('/api/header-aliases');
    state.headerAliases = data.aliases || {};
  } catch (_err) {
    state.headerAliases = {};
  }
}

async function initApp() {
  resetDocumentHorizontalScroll();
  window.addEventListener('pageshow', resetDocumentHorizontalScroll);
  window.addEventListener('resize', resetDocumentHorizontalScroll);
  bindEvents();
  autosaveStatus('Tự động lưu đang bật', 'ok');
  await loadHeaderAliases();
  await loadStaffList();
  await loadClsList();
  await loadFiles();
  await loadTheoSoFiles();
  await restoreBestDraftIfAny(true);
  updateContextUI();
  updateDashboardMetrics();
}

initApp().catch(err => {
  alert(err.message);
  log(`Lỗi khởi động: ${err.message}`);
  setBusy(false);
});
