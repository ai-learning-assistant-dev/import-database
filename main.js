// 临时兼容：Electron 所带 Node 版本可能未提供全局 File，undici 在 fetch 初始化时访问 File 导致报错。
// 提前定义一个简易 File polyfill 避免 "ReferenceError: File is not defined"。
if (typeof File === 'undefined') {
	global.File = class FilePolyfill {
		constructor(bits, name, options = {}) {
			this.name = name || 'unnamed';
			this.type = options.type || '';
			this.lastModified = options.lastModified || Date.now();
			// 估算 size（仅文本/Buffer 简单统计）
			if (Array.isArray(bits)) {
				this.size = bits.reduce((acc, b) => {
					if (typeof b === 'string') return acc + Buffer.byteLength(b);
					if (Buffer.isBuffer(b)) return acc + b.length;
					return acc;
				}, 0);
			} else {
				this.size = 0;
			}
		}
	};
}

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { Client } = require('pg');
const axios = require('axios');
const cheerio = require('cheerio');

// 开发模式主进程热重载（仅在 ELECTRON_DEV 环境变量为 1 时启用）
if (process.env.ELECTRON_DEV) {
	try {
		// 监听整个项目目录；忽略 node_modules 提升性能
		require('electron-reload')(__dirname, {
			electron: path.join(__dirname, 'node_modules', '.bin', 'electron'),
			ignored: /node_modules|[\\/]\./
		});
		console.log('[hot] electron-reload 已启用');
	} catch (e) {
		console.warn('[hot] electron-reload 初始化失败:', e.message);
	}
}

function getConfigFilePath() {
	return path.join(app.getPath('userData'), 'dbconfig.json');
}

function createWindow() {
	const win = new BrowserWindow({
		width: 900,
		height: 600,
		minWidth: 600,
		minHeight: 400,
		backgroundColor: '#000000e3',
		title: '课程内容导入工具',
		autoHideMenuBar: true,
		webPreferences: {
			contextIsolation: true,
			preload: path.join(__dirname, 'preload.js'),
			sandbox: false
		}
	});

	if (process.env.ELECTRON_DEV) {
		// 开发模式：加载 Vite Dev Server
		const devUrl = process.env.ELECTRON_START_URL || 'http://localhost:5173';
		console.log('[main] load dev url:', devUrl);
		win.loadURL(devUrl);
		win.webContents.openDevTools({ mode: 'detach' });
	} else {
		// 生产模式：优先加载 dist/index.html，不存在则回退到根 index.html 方便开发
		const distIndex = path.join(__dirname, 'dist', 'index.html');
		if (fs.existsSync(distIndex)) {
			console.log('[main] load dist index:', distIndex);
			win.loadFile(distIndex);
		} else {
			const fallback = path.join(__dirname, 'index.html');
			console.log('[main] fallback load root index:', fallback);
			win.loadFile(fallback);
		}
	}

	win.webContents.on('did-fail-load', (_e, errorCode, errorDescription, validatedURL) => {
		console.error('[main] did-fail-load', { errorCode, errorDescription, validatedURL });
		win.webContents.executeJavaScript(`document.body.innerHTML = '<pre style="color:#f55;padding:16px;font-family:monospace;">Load Failed: ${errorCode} ${errorDescription}<br>${validatedURL}</pre>';`);
	});

	win.webContents.on('did-finish-load', () => {
		console.log('[main] did-finish-load');
	});
}

// IPC handlers
ipcMain.handle('parse-excel', async (_event, filePath) => {
	try {
		const XLSX = require('xlsx');
		if (!fs.existsSync(filePath)) {
			return { ok: false, error: '文件不存在' };
		}
		const workbook = XLSX.readFile(filePath);
		if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
			return { ok: true, list: [] };
		}
		const sheetName = workbook.SheetNames[1];//为第二页
		const sheet = workbook.Sheets["chapters_sections"];
		// header:1 得到二维数组, 每行是一个数组, 便于在没有标题行时处理
		const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
		if (!rows.length) return { ok: true, list: [] };
		console.log(`[parse-excel] rows count: ${rows.length}`);
		const firstRow = rows[0].map(v => String(v).trim());
		// 统一用小写关键字匹配，以避免大小写不一致导致未识别
		const headerKeywordsUrl = ['url','链接','link','地址'];
		const headerKeywordsName = ['name','title','视频名称','名称'];
		const lowerFirst = firstRow.map(v => v.toLowerCase());
		let urlCol = -1;
		let nameCol = -1;
		lowerFirst.forEach((cell, idx) => {
			if (urlCol === -1 && headerKeywordsUrl.includes(cell)) urlCol = idx;
			if (nameCol === -1 && headerKeywordsName.includes(cell)) nameCol = idx;
		});
		// 如果没找到标题行中的 url 列, 尝试通过内容自动识别: 找到出现 http 开头次数最多的列
		if (urlCol === -1) {
			const httpCounts = new Map();
			for (let r = 0; r < rows.length; r++) {
				const row = rows[r];
				row.forEach((cell, cIdx) => {
					if (typeof cell === 'string' && cell.trim().startsWith('http')) {
						httpCounts.set(cIdx, (httpCounts.get(cIdx) || 0) + 1);
					}
				});
			}
			let maxCount = 0; let maxCol = -1;
			for (const [colIdx, count] of httpCounts.entries()) {
				if (count > maxCount) { maxCount = count; maxCol = colIdx; }
			}
			if (maxCount > 0) urlCol = maxCol;
		}
		// 如果没找到名称列, 尝试找一个非 URL 且文本较多的列
		if (nameCol === -1 && urlCol !== -1) {
			let candidateCol = -1; let candidateScore = -1;
			for (let c = 0; c < firstRow.length; c++) {
				if (c === urlCol) continue;
				let score = 0;
				for (let r = 1; r < rows.length; r++) {
					const val = rows[r][c];
					if (typeof val === 'string') {
						const trimmed = val.trim();
						if (trimmed && !trimmed.startsWith('http')) score += trimmed.length;
					}
				}
				if (score > candidateScore) { candidateScore = score; candidateCol = c; }
			}
			if (candidateScore > 0) nameCol = candidateCol;
		}
		// 数据起始行: 如果第一行包含任何 header 关键词则跳过它
		const headerPresent = lowerFirst.some(cell => headerKeywordsUrl.includes(cell) || headerKeywordsName.includes(cell));
		const startRow = headerPresent ? 1 : 0;
		const list = [];
		for (let r = startRow; r < rows.length; r++) {
			const rowArr = rows[r];
			const url = urlCol >= 0 ? String(rowArr[urlCol] || '').trim() : '';
			if (!url || !url.startsWith('http')) continue;
			const name = nameCol >= 0 ? String(rowArr[nameCol] || '').trim() : '';
			// 根据是否存在表头构建行对象：有表头则以表头文字为键，否则 col{index}
			const rowObj = {};
			if (headerPresent) {
				firstRow.forEach((h, idx) => { const key = h || `col${idx}`; rowObj[key] = rowArr[idx]; });
			} else {
				rowArr.forEach((val, idx) => { rowObj[`col${idx}`] = val; });
			}
			// 展开列名到返回对象，附加 url/name/rowIndex
			list.push({ ...rowObj, url, name, rowIndex: r });
		}

		return { ok: true, list };
	} catch (e) {
		return { ok: false, error: e.message };
	}
});
ipcMain.handle('get-time', () => {
	return new Date().toLocaleString();
});

ipcMain.handle('ping', (_event, msg) => {
	return `pong: ${msg}`;
});
ipcMain.handle('fetch-video-titles', async (_event, urls) => {
	if (!Array.isArray(urls)) return [];
	const tasks = urls.map(async (rawUrl) => {
		let url = typeof rawUrl === 'string' ? rawUrl.trim() : '';
		if (!url) return { url, title: '', imageSrc: '', ok: false, error: '空URL' };
		try {
			// 桌面版抓取标题
			const desktopRes = await axios.get(url, {
				headers: {
					'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
					'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
					'Accept-Language': 'zh-CN,zh;q=0.9'
				},
				timeout: 12000
			});
			const $d = cheerio.load(desktopRes.data);
			const titleAttr = $d('title').attr('data-title');
			const plainH1 = $d('h1').first().text().trim();
			const htmlTitle = $d('title').first().text().trim().replaceAll('_哔哩哔哩_bilibili','');
			const title = htmlTitle || (titleAttr && titleAttr.trim()) || plainH1 || '';

			let imageSrc = '';
			// 1. 优先 meta og:image / twitter:image 等静态封面
			imageSrc = $d('meta[property="og:image"]').attr('content')
				|| $d('meta[name="twitter:image"]').attr('content')
				|| '';
			if (imageSrc && imageSrc.startsWith('//')) imageSrc = 'https:' + imageSrc;

			// 2. 若为空并且是 bilibili 视频页面，尝试调用官方 API 获取封面
			if (!imageSrc && /https:\/\/www\.bilibili\.com\/.+\/video\//.test(url)) {
				const bvidMatch = url.match(/BV[0-9A-Za-z]+/);
				if (bvidMatch) {
					try {
						const apiRes = await axios.get(`https://api.bilibili.com/x/web-interface/view?bvid=${bvidMatch[0]}`, {
							headers: { 'User-Agent': 'Mozilla/5.0' }, timeout: 10000
						});
						if (apiRes.data && apiRes.data.data && apiRes.data.data.pic) {
							imageSrc = apiRes.data.data.pic;
							if (imageSrc.startsWith('//')) imageSrc = 'https:' + imageSrc;
						}
					} catch (_) {}
				}
			}

			// 如果是 bilibili 桌面链接，转换为移动端链接
			let mobileUrl = url;
			if (url.startsWith('https://www.bilibili.com')) {
				mobileUrl = url.replace('https://www.bilibili.com', 'https://m.bilibili.com');
			}

			// 3. 移动端补充：若仍无封面则尝试根据 alt 匹配 img
			try {
				const mobileRes = await axios.get(mobileUrl, {
					headers: {
						'User-Agent': 'Mozilla/5.0 (Linux; Android 10; Mobile) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36',
						'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
						'Accept-Language': 'zh-CN,zh;q=0.9'
					},
					timeout: 12000
				});
				const $m = cheerio.load(mobileRes.data);
				if (!imageSrc && title) {
					let img = $m(`img[alt="${title}"]`).first();
					if ((!img || !img.attr('src')) && title) {
						img = $m('img').filter((i, el) => $m(el).attr('alt') === title).first();
					}
					const mobileSrc = img && (img.attr('src') || img.attr('data-src') || '').trim();
					if (mobileSrc) {
						imageSrc = mobileSrc.startsWith('//') ? 'https:' + mobileSrc : mobileSrc;
					}
				}
			}
			catch (e2) {
				// 不影响标题抓取
			}
			return { url, title, imageSrc, ok: true };
		} catch (e) {
			return { url, title: '', imageSrc: '', ok: false, error: e.message };
		}
	});
	const results = await Promise.all(tasks);
	return results;
});

// 代理拉取外部图片并转为 base64，便于在渲染进程中使用
ipcMain.handle('fetch-image-base64', async (_event, url) => {
	try {
		if (!url) return { ok: false, error: '缺少 url' };
		const res = await axios.get(url, {
			responseType: 'arraybuffer',
			headers: {
				'User-Agent': 'Mozilla/5.0',
				'Referer': 'https://www.bilibili.com/'
			}
		});
		const contentType = res.headers['content-type'] || 'image/jpeg';
		const base64 = Buffer.from(res.data).toString('base64');
		return { ok: true, src: `data:${contentType};base64,${base64}` };
	} catch (e) {
		return { ok: false, error: e.message };
	}
});

ipcMain.handle('load-db-config', () => {
	try {
		const p = getConfigFilePath();
		if (fs.existsSync(p)) {
			const raw = fs.readFileSync(p, 'utf-8');
			return JSON.parse(raw);
		}
		return null;
	} catch (e) {
		return { error: '读取配置失败', detail: e.message };
	}
});

ipcMain.handle('save-db-config', (_event, cfg) => {
	try {
		const saveObj = {
			host: cfg.host || 'localhost',
			port: Number(cfg.port) || 5432,
			user: cfg.user || '',
			password: cfg.password || '',
			database: cfg.database || '',
			savedAt: new Date().toISOString()
		};
		fs.mkdirSync(app.getPath('userData'), { recursive: true });
		fs.writeFileSync(getConfigFilePath(), JSON.stringify(saveObj, null, 2), 'utf-8');
		return { ok: true };
	} catch (e) {
		return { ok: false, error: e.message };
	}
});

ipcMain.handle('test-db-connection', async (_event, cfg) => {
	const client = new Client({
		host: cfg.host,
		port: Number(cfg.port) || 5432,
		user: cfg.user,
		password: cfg.password,
		database: cfg.database,
		connectionTimeoutMillis: 5000
	});
	try {
		await client.connect();
		const res = await client.query('SELECT 1 AS alive');
		await client.end();
		return { ok: true, result: res.rows };
	} catch (e) {
		try { await client.end(); } catch (_) {}
		return { ok: false, error: e.message };
	}
});

app.whenReady().then(() => {
	createWindow();
	app.on('activate', () => {
		if (BrowserWindow.getAllWindows().length === 0) {
			createWindow();
		}
	});
});

app.on('window-all-closed', () => {
	if (process.platform !== 'darwin') {
		app.quit();
	}
});
// 选择文件（Excel）
ipcMain.handle('choose-file', async () => {
	const win = BrowserWindow.getFocusedWindow();
	const res = await dialog.showOpenDialog(win, {
		title: '选择 Excel 文件',
		properties: ['openFile'],
		filters: [
			{ name: 'Excel', extensions: ['xlsx', 'xls'] },
			{ name: '所有文件', extensions: ['*'] }
		]
	});
	if (res.canceled || res.filePaths.length === 0) {
		return null;
	}
	return res.filePaths[0];
});

// 选择素材文件夹（用于导入 sections）
ipcMain.handle('choose-folder', async () => {
	const win = BrowserWindow.getFocusedWindow();
	const res = await dialog.showOpenDialog(win, {
		title: '选择素材文件夹',
		properties: ['openDirectory']
	});
	if (res.canceled || res.filePaths.length === 0) return null;
	return res.filePaths[0];
});

// 查询 courses 表
ipcMain.handle('query-courses', async (_event, opts) => {
	const search = opts && typeof opts.search === 'string' ? opts.search.trim() : '';
	const limit = opts && Number(opts.limit) > 0 ? Math.min(Number(opts.limit), 200) : 50;
	// 读取已保存的数据库配置
	try {
		const cfgPath = getConfigFilePath();
		if (!fs.existsSync(cfgPath)) return { ok: false, error: '未配置数据库' };
		const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf-8'));
		const client = new Client({
			host: cfg.host,
			port: Number(cfg.port) || 5432,
			user: cfg.user,
			password: cfg.password,
			database: cfg.database,
			connectionTimeoutMillis: 5000
		});
		await client.connect();
		let sql = 'SELECT course_id, name, icon_url, description, default_ai_persona_id FROM courses';
		const params = [];
		if (search) {
			sql += ' WHERE name ILIKE $1 OR course_id::text = $1';
			params.push('%' + search + '%');
		}
		sql += ' ORDER BY name ASC LIMIT ' + limit;
		const res = await client.query(sql, params);
		await client.end();
		return { ok: true, list: res.rows };
	} catch (e) {
		return { ok: false, error: e.message };
	}
});

// 新增 / 更新课程（按名称 upsert）
ipcMain.handle('upsert-course', async (_event, payload) => {
	const name = payload && typeof payload.name === 'string' ? payload.name.trim() : '';
	const icon_url = payload && typeof payload.icon_url === 'string' ? payload.icon_url.trim() : '';
	const description = payload && typeof payload.description === 'string' ? payload.description.trim() : '';
	const default_ai_persona_id = payload && (payload.default_ai_persona_id === null || payload.default_ai_persona_id === undefined || payload.default_ai_persona_id === '' ? null : Number(payload.default_ai_persona_id));
	if (!name) return { ok: false, error: '名称必填' };
	if (!icon_url) return { ok: false, error: '图片( icon_url )必填' };
	try {
		const cfgPath = getConfigFilePath();
		if (!fs.existsSync(cfgPath)) return { ok: false, error: '未配置数据库' };
		const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf-8'));
		const client = new Client({
			host: cfg.host,
			port: Number(cfg.port) || 5432,
			user: cfg.user,
			password: cfg.password,
			database: cfg.database,
			connectionTimeoutMillis: 5000
		});
		await client.connect();
		// 查是否存在同名
		const existRes = await client.query('SELECT course_id, name FROM courses WHERE name = $1 LIMIT 1', [name]);
		let action = 'insert';
		let courseRow;
		if (existRes.rows.length) {
			// update
			action = 'update';
			const courseId = existRes.rows[0].course_id;
			await client.query(
				'UPDATE courses SET icon_url = $1, description = $2, default_ai_persona_id = $3 WHERE course_id = $4',
				[icon_url, description || null, default_ai_persona_id, courseId]
			);
			const fetchRes = await client.query('SELECT course_id, name, icon_url, description, default_ai_persona_id FROM courses WHERE course_id = $1', [courseId]);
			courseRow = fetchRes.rows[0];
		} else {
			// insert
			const insRes = await client.query(
				'INSERT INTO courses (name, icon_url, description, default_ai_persona_id) VALUES ($1,$2,$3,$4) RETURNING course_id, name, icon_url, description, default_ai_persona_id',
				[name, icon_url, description || null, default_ai_persona_id]
			);
			courseRow = insRes.rows[0];
		}
		await client.end();
		return { ok: true, action, course: courseRow };
	} catch (e) {
		return { ok: false, error: e.message };
	}
});

// 批量导入章节：去重章标题
ipcMain.handle('import-chapters', async (_event, payload) => {
	const courseId = payload && typeof payload.courseId === 'string' ? payload.courseId.trim() : '';
	const chapters = Array.isArray(payload && payload.chapters) ? payload.chapters : [];
	if (!courseId) return { ok: false, error: '缺少 courseId' };
	if (!chapters.length) return { ok: false, error: '章节列表为空' };
	try {
		const cfgPath = getConfigFilePath();
		if (!fs.existsSync(cfgPath)) return { ok: false, error: '未配置数据库' };
		const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf-8'));
		const client = new Client({
			host: cfg.host,
			port: Number(cfg.port) || 5432,
			user: cfg.user,
			password: cfg.password,
			database: cfg.database,
			connectionTimeoutMillis: 5000
		});
		await client.connect();
		// 读取已存在的标题
		const existRes = await client.query('SELECT title FROM chapters WHERE course_id = $1', [courseId]);
		const existingSet = new Set(existRes.rows.map(r => (r.title || '').trim()));
		// 去重传入章节
		const unique = [];
		const seen = new Set();
		for (const ch of chapters) {
			if (!ch) continue;
			const title = (ch.title || '').trim();
			if (!title) continue;
			if (seen.has(title)) continue; // 本次重复
			seen.add(title);
			if (existingSet.has(title)) continue; // 已存在
			const order = Number(ch.chapter_order);
			unique.push({ title, chapter_order: Number.isFinite(order) ? order : unique.length + 1 });
		}
		if (!unique.length) {
			await client.end();
			return { ok: true, inserted: 0, skipped: chapters.length };
		}
		// 批量插入
		const values = [];
		const params = [];
		let paramIndex = 1;
		for (const row of unique) {
			// (course_id, title, chapter_order)
			params.push(courseId, row.title, row.chapter_order);
			values.push(`($${paramIndex}, $${paramIndex + 1}, $${paramIndex + 2})`);
			paramIndex += 3;
		}
		const insertSql = `INSERT INTO chapters (course_id, title, chapter_order) VALUES ${values.join(', ')}`;
		await client.query(insertSql, params);
		await client.end();
		return { ok: true, inserted: unique.length, skipped: chapters.length - unique.length };
	} catch (e) {
		return { ok: false, error: e.message };
	}
});

// 清空所有 sections 记录
ipcMain.handle('clear-sections', async () => {
	try {
		const cfgPath = getConfigFilePath();
		if (!fs.existsSync(cfgPath)) return { ok: false, error: '未配置数据库' };
		const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf-8'));
		const client = new Client({
			host: cfg.host,
			port: Number(cfg.port) || 5432,
			user: cfg.user,
			password: cfg.password,
			database: cfg.database,
			connectionTimeoutMillis: 8000
		});
		await client.connect();
		// 先清空 leading_question，再清空 sections 表，保持外键约束安全
		await client.query('DELETE FROM leading_question');
		const res = await client.query('DELETE FROM sections');
		await client.end();
		return { ok: true, deleted: res.rowCount };
	} catch (e) {
		return { ok: false, error: e.message };
	}
});

// 导入 sections：依据 Excel 行与素材文件夹
ipcMain.handle('import-sections', async (_event, payload) => {
	const courseId = payload && typeof payload.courseId === 'string' ? payload.courseId.trim() : '';
	const rows = Array.isArray(payload && payload.rows) ? payload.rows : [];
	const folderPath = payload && typeof payload.folderPath === 'string' ? payload.folderPath.trim() : '';
	if (!courseId) return { ok: false, error: '缺少 courseId' };
	if (!rows.length) return { ok: false, error: 'Excel 行为空' };
	if (!folderPath) return { ok: false, error: '未选择素材文件夹' };
	try {
		const cfgPath = getConfigFilePath();
		if (!fs.existsSync(cfgPath)) return { ok: false, error: '未配置数据库' };
		const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf-8'));
		const client = new Client({
			host: cfg.host,
			port: Number(cfg.port) || 5432,
			user: cfg.user,
			password: cfg.password,
			database: cfg.database,
			connectionTimeoutMillis: 8000
		});
		await client.connect();
		// 章节映射
		const chapRes = await client.query('SELECT chapter_id, title FROM chapters WHERE course_id = $1', [courseId]);
		const chapterMap = new Map(chapRes.rows.map(r => [ (r.title || '').trim(), r.chapter_id ]));
		// 已有节集合
		const existRes = await client.query('SELECT s.title, s.chapter_id FROM sections s JOIN chapters c ON s.chapter_id = c.chapter_id WHERE c.course_id = $1', [courseId]);
		const existingKey = new Set(existRes.rows.map(r => r.chapter_id + '||' + (r.title || '').trim()));
		const sanitize = (name) => name.replace(/[\\/:*?"<>| ]/g, '_');
		const parseDuration = (val) => {
			if (val == null) return null;
			if (typeof val === 'number') return val;
			const s = String(val).trim();
			if (!s) return null;
			if (/^\d{1,2}:\d{1,2}(:\d{1,2})?$/.test(s)) {
				const parts = s.split(':').map(p => parseInt(p,10));
				while (parts.length < 3) parts.unshift(0);
				const [hh,mm,ss] = parts;
				return hh*3600 + mm*60 + ss;
			}
			const num = parseFloat(s);
			if (!Number.isFinite(num)) return null;
			// 认为数值字段本身就是秒数
			return num;
		};
		const parseSrtToJson = (content) => {
			const lines = content.replace(/\r/g,'').split('\n');
			const entries = [];
			let i=0;
			while (i < lines.length) {
				const idxLine = lines[i].trim();
				if (!idxLine) { i++; continue; }
				const seq = parseInt(idxLine,10);
				if (!Number.isFinite(seq)) { i++; continue; }
				i++;
				if (i >= lines.length) break;
				const timeLine = lines[i].trim();
				const m = timeLine.match(/^(\d{2}:\d{2}:\d{2},\d{3}) --> (\d{2}:\d{2}:\d{2},\d{3})/);
				if (!m) { i++; continue; }
				const start = m[1];
				const end = m[2];
				i++;
				const textLines = [];
				while (i < lines.length && lines[i].trim() !== '') {
					textLines.push(lines[i]);
					i++;
				}
				while (i < lines.length && lines[i].trim() === '') i++;
				entries.push({ seq, start, end, text: textLines.join('\n') });
			}
			return entries;
		};
		let insertedCount = 0;
		let skipped = 0;
		const missingFiles = [];
		for (const row of rows) {
			const chapterTitle = (row['章标题'] || '').trim();
			const sectionTitle = (row['节标题'] || '').trim();
			if (!chapterTitle || !sectionTitle) { skipped++; continue; }
			const chapterId = chapterMap.get(chapterTitle);
			if (!chapterId) { missingFiles.push({ sectionTitle, reason:'章未找到:'+chapterTitle }); skipped++; continue; }
			const key = chapterId + '||' + sectionTitle;
			if (existingKey.has(key)) { skipped++; continue; }
			const scrapedTitle = (row['title'] || row.title || '').trim() || sectionTitle;
			const baseCandidates = [scrapedTitle];
			let srtFileRel = null; // 记录找到的字幕文件（相对素材目录名）
			let summaryJson = null, subtitlesJson = null, markdownContent = null, questionsList = [];
			let exercisesJson = null; // _exercises.json 解析结果
			let foundSummary=false, foundSrt=false, foundMd=false;
			for (const base of baseCandidates) {
				if (!base) continue;
				const bn = sanitize(base);
				const summaryPath = path.join(folderPath, bn.trim() + '_summary.json');
				if (!foundSummary && fs.existsSync(summaryPath)) {
					 try { summaryJson = JSON.parse(fs.readFileSync(summaryPath,'utf-8'));
						 foundSummary=true; } 
						 catch (e){
						 missingFiles.push({ sectionTitle, reason:'解析summary失败:'+bn });
					} }
				const srtPath = path.join(folderPath, bn + '_ai-zh.srt');
				if (!foundSrt && fs.existsSync(srtPath)) {
					try {
						const raw = fs.readFileSync(srtPath,'utf-8');
						subtitlesJson = parseSrtToJson(raw);
						foundSrt = true;
						srtFileRel = bn + '_ai-zh.srt';
					} catch (e){
						missingFiles.push({ sectionTitle, reason:'解析srt失败:'+bn });
					}
				}
				const mdPath = path.join(folderPath, bn + '.md');
				if (!foundMd && fs.existsSync(mdPath)) { try { markdownContent = fs.readFileSync(mdPath,'utf-8'); foundMd=true; } catch (e){ missingFiles.push({ sectionTitle, reason:'读取md失败:'+bn }); } }
				// questions.json
				if (!questionsList || questionsList.length === 0) {
					const qPath = path.join(folderPath, bn + '_questions.json');
					if (fs.existsSync(qPath)) {
						try {
							const rawQ = fs.readFileSync(qPath, 'utf-8');
							const qJson = JSON.parse(rawQ);
							if (qJson && Array.isArray(qJson.questions)) {
								questionsList = qJson.questions
									.map(item => item && item.question && String(item.question).trim())
									.filter(Boolean);
							}
						} catch (e) {
							missingFiles.push({ sectionTitle, reason:'解析questions失败:'+bn });
						}
					}
				}
				// exercises.json
				if (!exercisesJson) {
					const exPath = path.join(folderPath, bn + '_exercises.json');
					if (fs.existsSync(exPath)) {
						try {
							const rawEx = fs.readFileSync(exPath, 'utf-8');
							exercisesJson = JSON.parse(rawEx);
						} catch (e) {
							missingFiles.push({ sectionTitle, reason:'解析exercises失败:'+bn });
						}
					}
				}
			}
			if (!foundSummary) missingFiles.push({ sectionTitle, reason:'summary缺失' });
			if (!foundSrt) missingFiles.push({ sectionTitle, reason:'srt缺失' });
			if (!foundMd) missingFiles.push({ sectionTitle, reason:'md缺失' });
			let videoUrl = (row['url'] || row['URL'] || '').trim();
			if (!videoUrl) videoUrl = null;
			const strPath = (row['imageSrc'] || '').trim() || null;
			let estimatedTime = null;
			const durationSecFromExcel = parseDuration(row['时长']);
			if (Number.isFinite(durationSecFromExcel) && durationSecFromExcel > 0) {
				estimatedTime = Math.ceil(durationSecFromExcel / 60);
			}
			// 解析 _video_info.json 优先覆盖  estimatedTime
			let foundVideoInfo = false;
			for (const base of baseCandidates) {
				if (foundVideoInfo) break;
				if (!base) continue;
				const bn = sanitize(base);
				const infoPath = path.join(folderPath, bn + '_video_info.json');
				if (fs.existsSync(infoPath)) {
					try {
						const raw = fs.readFileSync(infoPath, 'utf-8');
						const info = JSON.parse(raw);
						foundVideoInfo = true;
						if (info && info.duration !== undefined && info.duration !== null) {
							let durationSec = null;
							if (typeof info.duration === 'number' && Number.isFinite(info.duration)) {
								durationSec = info.duration;
							} else if (typeof info.duration === 'string') {
								const ds = info.duration.trim();
								if (/^\d+(\.\d+)?$/.test(ds)) {
									durationSec = Number(ds);
								} else {
									const parsed = parseDuration(ds);
									if (Number.isFinite(parsed)) durationSec = parsed;
								}
							}
							if (Number.isFinite(durationSec) && durationSec > 0) {
								estimatedTime = Math.ceil(durationSec / 60);
							}
						}
					} catch (_) {
						// 忽略解析错误，保留 Excel 原值
					}
				}
			}
			let sectionOrder = parseInt(row['节顺序'],10);
			if (!Number.isFinite(sectionOrder)) sectionOrder = null;
			// 逐条插入 section，并立刻插入对应的 leading_question
			const secSql = 'INSERT INTO sections (title, chapter_id, video_url, knowledge_points, video_subtitles,srt_path, knowledge_content, estimated_time, section_order) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9) RETURNING section_id';
			const secParams = [
				sectionTitle,
				chapterId,
				videoUrl || null,
				summaryJson ? JSON.stringify(summaryJson) : null,
				subtitlesJson ? JSON.stringify(subtitlesJson) : null,
				strPath,
				markdownContent,
				estimatedTime,
				sectionOrder
			];
			const secRes = await client.query(secSql, secParams);
			const sectionIdInserted = secRes.rows[0] && secRes.rows[0].section_id;
			if (sectionIdInserted && Array.isArray(questionsList) && questionsList.length) {
				const qValues = [];
				const qParams = [];
				let qIdx = 1;
				for (const q of questionsList) {
					if (!q || !String(q).trim()) continue;
					qValues.push(`($${qIdx++}, $${qIdx++})`);
					qParams.push(sectionIdInserted, String(q).trim());
				}
				if (qValues.length) {
					const qSql = 'INSERT INTO leading_question (section_id, question) VALUES ' + qValues.join(',');
					await client.query(qSql, qParams);
				}
			}
			// 处理 exercises.json：插入 exercises 与 exercise_options
			if (sectionIdInserted && exercisesJson) {
				// multiple_choice 题目
				if (Array.isArray(exercisesJson.multiple_choice)) {
					for (const item of exercisesJson.multiple_choice) {
						if (!item || !item.question) continue;
						const exQuestion = String(item.question).trim();
						if (!exQuestion) continue;
						const exTypeStatus = '0'; // 单选
						const exScore = 10;// 默认10分
						const exAnswer = null;
						const exImage = null;
						const exSql = 'INSERT INTO exercises (section_id, question, type_status, score, answer, image) VALUES ($1,$2,$3,$4,$5,$6) RETURNING exercise_id';
						const exParams = [sectionIdInserted, exQuestion, exTypeStatus, exScore, exAnswer, exImage];
						const exRes = await client.query(exSql, exParams);
						const exerciseIdInserted = exRes.rows[0] && exRes.rows[0].exercise_id;
						// 插入 options
						if (exerciseIdInserted && item.options && typeof item.options === 'object') {
							const correctKey = typeof item.correct_answer === 'string' ? item.correct_answer.trim() : '';
							const optEntries = Object.entries(item.options);
							if (optEntries.length) {
								const optValues = [];
								const optParams = [];
								let oIdx = 1;
								for (const [keyOpt, textOpt] of optEntries) {
									const optText = textOpt && String(textOpt).trim();
									if (!optText) continue;
									const isCorrect = keyOpt === correctKey;
									optValues.push(`($${oIdx++}, $${oIdx++}, $${oIdx++}, $${oIdx++})`);
									optParams.push(exerciseIdInserted, optText, isCorrect, null);
								}
								if (optValues.length) {
									const optSql = 'INSERT INTO exercise_options (exercise_id, option_text, is_correct, image) VALUES ' + optValues.join(',');
									await client.query(optSql, optParams);
								}
							}
						}
					}
				}
				// short_answer 题目
				if (Array.isArray(exercisesJson.short_answer)) {
					for (const item of exercisesJson.short_answer) {
						if (!item || !item.question) continue;
						const exQuestion = String(item.question).trim();
						if (!exQuestion) continue;
						const exTypeStatus = '2'; // 简答
						const exScore = 10;
						const exAnswer = item.reference_answer ? String(item.reference_answer).trim() : null;
						const exImage = null;
						const exSql = 'INSERT INTO exercises (section_id, question, type_status, score, answer, image) VALUES ($1,$2,$3,$4,$5,$6)';
						const exParams = [sectionIdInserted, exQuestion, exTypeStatus, exScore, exAnswer, exImage];
						await client.query(exSql, exParams);
					}
				}
			}
			insertedCount++;
		}
		await client.end();
		return { ok:true, inserted: insertedCount, skipped, missingFiles };
	} catch (e) {
		return { ok:false, error: e.message };
	}
});

// 按课程级联清空：删除该课程下的 chapters、sections、leading_question、exercises、exercise_options
ipcMain.handle('clear-course-data', async (_event, payload) => {
	const rawCourseId = payload && payload.courseId;
	const courseId = rawCourseId != null ? String(rawCourseId).trim() : '';
	if (!courseId) return { ok: false, error: '缺少 courseId' };
	let client;
	try {
		const cfgPath = getConfigFilePath();
		if (!fs.existsSync(cfgPath)) return { ok: false, error: '未配置数据库' };
		const cfg = JSON.parse(fs.readFileSync(cfgPath, 'utf-8'));
		client = new Client({
			host: cfg.host,
			port: Number(cfg.port) || 5432,
			user: cfg.user,
			password: cfg.password,
			database: cfg.database,
			connectionTimeoutMillis: 8000
		});
		await client.connect();
		await client.query('BEGIN');
		// 找出该课程下所有 section_id
		const secRes = await client.query(
			`SELECT s.section_id FROM sections s
			 JOIN chapters c ON s.chapter_id = c.chapter_id
			 WHERE c.course_id = $1`,
			[courseId]
		);
		const sectionIds = secRes.rows.map(r => r.section_id);
		let deletedExerciseOptions = 0;
		let deletedExercises = 0;
		let deletedLeading = 0;
		let deletedSections = 0;
		let deletedChapters = 0;
		if (sectionIds.length) {
			// 删除练习题选项
			const exRes = await client.query('SELECT exercise_id FROM exercises WHERE section_id = ANY($1::uuid[])', [sectionIds]);
			const exerciseIds = exRes.rows.map(r => r.exercise_id);
			if (exerciseIds.length) {
				const delOpt = await client.query('DELETE FROM exercise_options WHERE exercise_id = ANY($1::uuid[])', [exerciseIds]);
				deletedExerciseOptions = delOpt.rowCount || 0;
			}
			const delEx = await client.query('DELETE FROM exercises WHERE section_id = ANY($1::uuid[])', [sectionIds]);
			deletedExercises = delEx.rowCount || 0;
			const delLead = await client.query('DELETE FROM leading_question WHERE section_id = ANY($1::uuid[])', [sectionIds]);
			deletedLeading = delLead.rowCount || 0;
			const delSec = await client.query('DELETE FROM sections WHERE section_id = ANY($1::uuid[])', [sectionIds]);
			deletedSections = delSec.rowCount || 0;
		}
		// 删除该课程的章节
		const delChap = await client.query('DELETE FROM chapters WHERE course_id = $1', [courseId]);
		deletedChapters = delChap.rowCount || 0;
		await client.query('COMMIT');
		await client.end();
		return {
			ok: true,
			deletedExerciseOptions,
			deletedExercises,
			deletedLeading,
			deletedSections,
			deletedChapters
		};
	} catch (e) {
		if (client) {
			try { await client.query('ROLLBACK'); } catch(_) {}
			try { await client.end(); } catch(_) {}
		}
		return { ok:false, error: e.message };
	}
});

// （已替换为上方版本）