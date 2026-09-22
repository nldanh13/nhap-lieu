const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

const APP_ROOT = __dirname;
const PORT = process.env.PORT || 3002;
const PYTHON = process.env.PYTHON || (process.platform === 'win32' ? 'python' : 'python3');
const uploadDir = path.join(APP_ROOT, 'node_uploads');
const outputDir = path.join(APP_ROOT, 'node_outputs');
const autosaveDir = path.join(APP_ROOT, 'node_autosaves');
const theoSoDir = path.join(APP_ROOT, 'TheoSo');
const autosaveStateFile = path.join(autosaveDir, 'autosave_state.json');
const emrConfigFile = path.join(APP_ROOT, 'emr_config.json');
const WORK_SUFFIX = '_NHAP_LIEU';
const GENERATED_SUFFIXES = [
  '_NHAP_LIEU', '_da_nhap_lieu_node', '_da_them_dong_node', '_da_xoa_dong_node',
  '_da_bo_sung_phu_mo_node', '_da_chuyen_tieu_phau_node', '_da_dien_BS_node',
  '_da_bo_sung_phu_mo', '_da_chuyen_tieu_phau', '_da_chuyen_tieu_phau_da_dien_BS',
  '_da_dien_BS'
];
fs.mkdirSync(uploadDir, { recursive: true });
fs.mkdirSync(outputDir, { recursive: true });
fs.mkdirSync(autosaveDir, { recursive: true });
fs.mkdirSync(theoSoDir, { recursive: true });

const app = express();
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true, limit: '20mb' }));
app.use(express.static(path.join(APP_ROOT, 'public')));

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, uploadDir),
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname || '.xlsx') || '.xlsx';
    const base = path.basename(file.originalname || 'upload', ext).replace(/[\\/:*?"<>|]/g, '_');
    cb(null, `${base}_${Date.now()}${ext}`);
  }
});
const upload = multer({ storage });

const theoSoStorage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, theoSoDir),
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname || '.xlsx') || '.xlsx';
    const base = path.basename(file.originalname || 'upload', ext).replace(/[\\/:*?"<>|]/g, '_');
    cb(null, `${base}${ext}`);
  }
});
const uploadTheoSo = multer({ storage: theoSoStorage });

function safeResolve(filePath) {
  if (!filePath) throw new Error('Thiếu đường dẫn file.');
  const resolved = path.isAbsolute(filePath)
    ? path.resolve(filePath)
    : path.resolve(APP_ROOT, filePath);
  const roots = [APP_ROOT, uploadDir, outputDir].map(p => path.resolve(p));
  if (!roots.some(root => resolved === root || resolved.startsWith(root + path.sep))) {
    throw new Error('Đường dẫn file không nằm trong thư mục chương trình.');
  }
  return resolved;
}

function relativeToApp(filePath) {
  const resolved = safeResolve(filePath);
  return path.relative(APP_ROOT, resolved);
}

function cleanWorkbookStem(name) {
  let stem = String(name || '').replace(/%20/g, ' ');
  let changed = true;
  while (changed) {
    changed = false;
    for (const suffix of GENERATED_SUFFIXES) {
      if (stem.endsWith(suffix)) {
        stem = stem.slice(0, -suffix.length);
        changed = true;
      }
    }
  }
  return stem.replace(/[\/:*?"<>|]/g, '_').trim() || 'PM_khoa_CTCH_T5';
}

function outputPathFor(inputFile, _suffix) {
  const resolved = safeResolve(inputFile);
  const parsed = path.parse(resolved);
  const ext = parsed.ext || '.xlsx';
  const cleanStem = cleanWorkbookStem(parsed.name);
  return path.join(outputDir, `${cleanStem}${WORK_SUFFIX}${ext}`);
}


function outputPathForUploadedName(originalName) {
  const parsed = path.parse(String(originalName || 'upload.xlsx'));
  const ext = ['.xlsx', '.xlsm'].includes(parsed.ext.toLowerCase()) ? parsed.ext : '.xlsx';
  const cleanStem = cleanWorkbookStem(parsed.name);
  return path.join(outputDir, `${cleanStem}${WORK_SUFFIX}${ext}`);
}

function runPython(args, options = {}) {
  return new Promise((resolve, reject) => {
    const pyArgs = [path.join(APP_ROOT, 'node_excel_api.py'), ...args];
    const child = spawn(PYTHON, pyArgs, {
      cwd: APP_ROOT,
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' }
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', d => { stdout += d.toString('utf8'); });
    child.stderr.on('data', d => { stderr += d.toString('utf8'); });
    child.on('error', reject);
    child.on('close', code => {
      const lines = stdout.trim().split(/\r?\n/).filter(Boolean);
      const lastLine = lines.length ? lines[lines.length - 1] : '';
      let parsed;
      try {
        parsed = JSON.parse(lastLine || '{}');
      } catch (err) {
        const message = stderr || stdout || err.message;
        reject(new Error(message));
        return;
      }
      if (code !== 0 || parsed.ok === false) {
        reject(new Error(parsed.error || stderr || `Python exited with code ${code}`));
        return;
      }
      parsed.rawStdout = options.includeRaw ? stdout : undefined;
      parsed.stderr = stderr || undefined;
      resolve(parsed);
    });
  });
}


function runJsonPython(scriptName, args = [], options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(PYTHON, [path.join(APP_ROOT, scriptName), ...args], {
      cwd: APP_ROOT,
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' }
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', d => { stdout += d.toString('utf8'); });
    child.stderr.on('data', d => { stderr += d.toString('utf8'); });
    child.on('error', reject);
    child.on('close', code => {
      const lines = stdout.trim().split(/\r?\n/).filter(Boolean);
      const lastLine = lines.length ? lines[lines.length - 1] : '';
      let parsed;
      try {
        parsed = JSON.parse(lastLine || '{}');
      } catch (err) {
        reject(new Error(stderr || stdout || err.message));
        return;
      }
      if (code !== 0 || parsed.ok === false) {
        reject(new Error(parsed.error || stderr || `Python exited with code ${code}`));
        return;
      }
      if (options.includeRaw) parsed.rawStdout = stdout;
      if (stderr) parsed.stderr = stderr;
      resolve(parsed);
    });
  });
}

function readEmrConfigRaw() {
  const defaults = {
    url_login: 'http://192.168.2.26:2026/login.aspx',
    username: '',
    password: '',
    login_user_field: 'txtLoginName',
    login_pass_field: 'txtPassword',
    login_button_field: 'btnLogin',
    page_load_timeout: 30,
    script_timeout: 30,
    page_load_strategy: 'eager',
    headless: true,
  };
  try {
    if (!fs.existsSync(emrConfigFile)) return defaults;
    const loaded = JSON.parse(fs.readFileSync(emrConfigFile, 'utf8') || '{}');
    return { ...defaults, ...(loaded || {}) };
  } catch (_err) {
    return defaults;
  }
}

function writeEmrConfig(input) {
  const current = readEmrConfigRaw();
  const next = {
    ...current,
    url_login: String(input.url_login || current.url_login || '').trim(),
    username: String(input.username || current.username || '').trim(),
    password: input.password === undefined || input.password === ''
      ? String(current.password || '')
      : String(input.password),
    headless: input.headless === undefined ? Boolean(current.headless) : Boolean(input.headless),
  };
  if (!next.url_login) throw new Error('Thiếu địa chỉ đăng nhập EMR.');
  if (!next.username) throw new Error('Thiếu tên đăng nhập EMR.');
  if (!next.password) throw new Error('Thiếu mật khẩu EMR.');
  const tmp = emrConfigFile + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(next, null, 2), 'utf8');
  fs.renameSync(tmp, emrConfigFile);
  return next;
}

function asyncHandler(fn) {
  return (req, res) => Promise.resolve(fn(req, res)).catch(err => {
    res.status(500).json({ ok: false, error: err.message || String(err) });
  });
}


function readAutosaveState() {
  try {
    if (!fs.existsSync(autosaveStateFile)) return {};
    const raw = fs.readFileSync(autosaveStateFile, 'utf8');
    const state = JSON.parse(raw || '{}');
    let resolvedFile = '';
    if (state.fileRelative) {
      resolvedFile = safeResolve(state.fileRelative);
    } else if (state.file) {
      // Tương thích file autosave cũ. Nếu đường dẫn cũ không còn tồn tại thì
      // bỏ trạng thái thay vì làm ứng dụng lỗi khi khởi động.
      resolvedFile = safeResolve(state.file);
    }
    if (resolvedFile && !fs.existsSync(resolvedFile)) return {};
    return {
      ...state,
      file: resolvedFile || '',
    };
  } catch (_err) {
    return {};
  }
}

function writeAutosaveState(data) {
  const source = { ...(data || {}) };
  let fileRelative = source.fileRelative || '';
  if (source.file) {
    fileRelative = relativeToApp(source.file);
  }
  delete source.file;
  const payload = {
    ...source,
    fileRelative,
    updatedAt: new Date().toISOString(),
  };
  fs.mkdirSync(autosaveDir, { recursive: true });
  const tmp = autosaveStateFile + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(payload, null, 2), 'utf8');
  fs.renameSync(tmp, autosaveStateFile);
  return payload;
}

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, appRoot: APP_ROOT, port: PORT, python: PYTHON });
});


app.get('/api/autosave-state', (_req, res) => {
  const state = readAutosaveState();
  res.json({ ok: true, state });
});

app.post('/api/autosave-state', asyncHandler(async (req, res) => {
  const body = req.body || {};
  if (body.file) safeResolve(body.file);
  const state = writeAutosaveState(body);
  res.json({ ok: true, state });
}));

app.post('/api/clear-autosave-state', (_req, res) => {
  try {
    if (fs.existsSync(autosaveStateFile)) fs.unlinkSync(autosaveStateFile);
  } catch (_err) {}
  res.json({ ok: true });
});

app.get('/api/staff-list', asyncHandler(async (_req, res) => {
  const data = await runPython(['staff-list']);
  res.json(data);
}));

app.post('/api/staff-list', asyncHandler(async (req, res) => {
  const staff = Array.isArray(req.body?.staff) ? req.body.staff : [];
  const data = await runPython(['save-staff-list', '--staff-json', JSON.stringify(staff)]);
  res.json(data);
}));

app.get('/api/cls-list', asyncHandler(async (_req, res) => {
  const data = await runPython(['cls-list']);
  res.json(data);
}));

app.post('/api/cls-list', asyncHandler(async (req, res) => {
  const cls = Array.isArray(req.body?.cls) ? req.body.cls : [];
  const data = await runPython(['save-cls-list', '--cls-json', JSON.stringify(cls)]);
  res.json(data);
}));

app.post('/api/finalize-workbook', asyncHandler(async (req, res) => {
  const file = safeResolve(req.body?.file);
  const data = await runPython(['finalize-workbook', '--file', file, '--output', file], { includeRaw: true });
  res.json(data);
}));


app.get('/api/emr-config', (_req, res) => {
  const config = readEmrConfigRaw();
  res.json({
    ok: true,
    config: {
      url_login: config.url_login || '',
      username: config.username || '',
      headless: Boolean(config.headless),
      configured: Boolean(config.url_login && config.username && config.password),
    }
  });
});

app.post('/api/emr-config', asyncHandler(async (req, res) => {
  const config = writeEmrConfig(req.body || {});
  res.json({
    ok: true,
    config: {
      url_login: config.url_login,
      username: config.username,
      headless: Boolean(config.headless),
      configured: true,
    }
  });
}));

app.post('/api/emr-fill-doctor', asyncHandler(async (req, res) => {
  const file = safeResolve(req.body.file);
  const sheet = String(req.body.sheet || '').trim();
  if (sheet.toLowerCase().replace(/\s+/g, '') !== 'tieuphau') {
    throw new Error('Chức năng EMR chỉ áp dụng cho sheet tieuphau.');
  }
  const config = readEmrConfigRaw();
  if (!config.url_login || !config.username || !config.password) {
    const error = new Error('Chưa cấu hình tài khoản EMR.');
    error.code = 'EMR_CONFIG_REQUIRED';
    throw error;
  }
  const isWorkFile = /_NHAP_LIEU\.(xlsx|xlsm)$/i.test(file);
  const output = isWorkFile ? file : outputPathFor(file);
  const args = [
    '--file', file,
    '--output', output,
    '--config', emrConfigFile,
  ];
  const rowNumber = Number(req.body.rowNumber || 0);
  if (rowNumber > 0) args.push('--row', String(rowNumber));
  if (config.headless) args.push('--headless');
  else args.push('--show-browser');
  const data = await runJsonPython('emr_tieuphau_fill.py', args, { includeRaw: true });
  res.json(data);
}));

app.get('/api/files', asyncHandler(async (_req, res) => {
  const data = await runPython(['list-files', '--root', APP_ROOT]);
  res.json(data);
}));

app.post('/api/upload', upload.single('file'), asyncHandler(async (req, res) => {
  if (!req.file) throw new Error('Chưa chọn file Excel.');
  const ext = path.extname(req.file.originalname || '').toLowerCase();
  if (!['.xlsx', '.xlsm'].includes(ext)) {
    try { fs.unlinkSync(req.file.path); } catch (_err) {}
    throw new Error('Chỉ hỗ trợ file Excel .xlsx hoặc .xlsm.');
  }

  // File upload được giữ nguyên trong node_uploads. Ngay sau khi nhận file,
  // chương trình tạo file làm việc và chuẩn hóa toàn bộ cột ngày/giờ. Nhờ vậy
  // số seri Excel như 46182 không còn xuất hiện trên bảng nhập liệu.
  const workFile = outputPathForUploadedName(req.file.originalname);
  const dateNormalization = await runPython([
    'normalize-dates', '--file', req.file.path, '--output', workFile
  ]);

  res.json({
    ok: true,
    file: dateNormalization.file || workFile,
    originalFile: req.file.path,
    name: req.file.originalname,
    storedName: req.file.filename,
    dateNormalization,
  });
}));

app.get('/api/theo-so/files', asyncHandler(async (_req, res) => {
  const files = fs.readdirSync(theoSoDir)
    .filter(name => /\.(xlsx|xls)$/i.test(name) && !name.startsWith('~$'))
    .map(name => {
      const stat = fs.statSync(path.join(theoSoDir, name));
      return { name, size: stat.size, modifiedAt: stat.mtime };
    })
    .sort((a, b) => a.name.localeCompare(b.name, 'vi'));
  res.json({ ok: true, files });
}));

app.post('/api/theo-so/upload', uploadTheoSo.array('files', 20), asyncHandler(async (req, res) => {
  const files = req.files || [];
  if (!files.length) throw new Error('Chưa chọn file Excel.');

  const invalid = files.filter(f => !['.xlsx', '.xls'].includes(path.extname(f.originalname || '').toLowerCase()));
  if (invalid.length) {
    for (const f of files) {
      try { fs.unlinkSync(f.path); } catch (_err) {}
    }
    throw new Error('Chỉ hỗ trợ file Excel .xlsx hoặc .xls (Sổ Thủ Thuật / Sổ Phẫu Thuật).');
  }

  res.json({ ok: true, uploaded: files.map(f => f.filename) });
}));

app.get('/api/workbook-info', asyncHandler(async (req, res) => {
  const file = safeResolve(req.query.file);
  const data = await runPython(['workbook-info', '--file', file]);
  res.json(data);
}));

app.get('/api/sheet-data', asyncHandler(async (req, res) => {
  const file = safeResolve(req.query.file);
  const args = ['read-sheet', '--file', file];
  if (req.query.sheet) args.push('--sheet', String(req.query.sheet));
  if (req.query.query) args.push('--query', String(req.query.query));
  if (String(req.query.missingOnly || '') === '1') args.push('--missing-only');
  args.push('--limit', String(req.query.limit ?? '0'));
  const data = await runPython(args);
  res.json(data);
}));

app.post('/api/update-row', asyncHandler(async (req, res) => {
  const file = safeResolve(req.body.file);
  const output = req.body.overwrite ? file : outputPathFor(file, req.body.suffix || '_da_nhap_lieu_node');
  const data = await runPython([
    'update-row', '--file', file, '--sheet', req.body.sheet, '--row', String(req.body.rowNumber),
    '--data-json', JSON.stringify(req.body.data || {}), '--output', output
  ]);
  res.json(data);
}));

app.post('/api/update-rows', asyncHandler(async (req, res) => {
  const file = safeResolve(req.body.file);
  const output = req.body.overwrite ? file : outputPathFor(file, req.body.suffix || '_NHAP_LIEU');
  const rows = Array.isArray(req.body.rows) ? req.body.rows : [];
  if (!rows.length) throw new Error('Không có dòng nào cần lưu.');
  const data = await runPython([
    'update-rows', '--file', file, '--sheet', req.body.sheet,
    '--rows-json', JSON.stringify(rows), '--output', output
  ]);
  res.json(data);
}));


app.post('/api/clear-columns', asyncHandler(async (req, res) => {
  const file = safeResolve(req.body.file);
  const columns = Array.isArray(req.body.columns) ? req.body.columns : [];
  if (!columns.length) throw new Error('Chưa chọn cột cần xóa.');
  const preview = Boolean(req.body.preview);
  const args = [
    'clear-columns', '--file', file, '--sheet', req.body.sheet,
    '--columns-json', JSON.stringify(columns)
  ];
  if (preview) {
    args.push('--preview');
  } else {
    const output = req.body.overwrite ? file : outputPathFor(file, req.body.suffix || '_NHAP_LIEU');
    args.push('--output', output);
  }
  const data = await runPython(args);
  res.json(data);
}));

app.post('/api/insert-row', asyncHandler(async (req, res) => {
  const file = safeResolve(req.body.file);
  const output = req.body.overwrite ? file : outputPathFor(file, req.body.suffix || '_da_them_dong_node');
  const args = [
    'insert-row', '--file', file, '--sheet', req.body.sheet,
    '--data-json', JSON.stringify(req.body.data || {}), '--output', output
  ];
  if (req.body.insertAt) args.push('--insert-at', String(req.body.insertAt));
  const data = await runPython(args);
  res.json(data);
}));

app.post('/api/delete-row', asyncHandler(async (req, res) => {
  const file = safeResolve(req.body.file);
  const output = req.body.overwrite ? file : outputPathFor(file, req.body.suffix || '_da_xoa_dong_node');
  const data = await runPython([
    'delete-row', '--file', file, '--sheet', req.body.sheet, '--row', String(req.body.rowNumber), '--output', output
  ]);
  res.json(data);
}));

app.post('/api/run-task', asyncHandler(async (req, res) => {
  const file = safeResolve(req.body.file);
  const task = String(req.body.task || '');
  const cmdMap = {
    'tat-ca': 'task-tat-ca',
    'bo-sung-phu-mo': 'task-bo-sung-phu-mo',
    'chuyen-tieu-phau': 'task-chuyen-tieu-phau',
    'cap-nhat-cls-tieuphau': 'task-cap-nhat-cls-tieuphau',
    'dien-bs': 'task-dien-bs'
  };
  if (!cmdMap[task]) throw new Error('Chức năng tự động không hợp lệ.');
  const output = outputPathFor(file);
  const data = await runPython([cmdMap[task], '--file', file, '--output', output], { includeRaw: true });
  res.json(data);
}));

app.get('/download', asyncHandler(async (req, res) => {
  const file = safeResolve(req.query.file);
  if (!fs.existsSync(file)) throw new Error('File không tồn tại.');
  if (String(req.query.finalize || '') === '1') {
    await runPython(['finalize-workbook', '--file', file, '--output', file], { includeRaw: true });
  }
  res.download(file);
}));

app.get('*', (_req, res) => {
  res.sendFile(path.join(APP_ROOT, 'public', 'index.html'));
});

app.listen(PORT, '127.0.0.1', () => {
  console.log('=========================================');
  console.log(`WEB APP NHẬP LIỆU ĐANG CHẠY`);
  console.log(`Mở trình duyệt: http://127.0.0.1:${PORT}`);
  console.log('Nhấn Ctrl + C để dừng server.');
  console.log('=========================================');
});
