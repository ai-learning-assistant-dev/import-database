# import-database

Electron + React + PostgreSQL 配置演示。支持通过安全桥接保存数据库连接信息并测试连接。

## 目录结构
- `main.js` 主进程：创建窗口、处理 IPC（时间、Ping、数据库配置保存与连接测试）。
- `preload.js` 预加载脚本：使用 `contextBridge` 暴露受限 API (`bridge`).
- `index.html` Vite 入口，挂载 React 应用。
- `src/main.jsx` React 组件：表单填写数据库配置并调用 `bridge` API。
- `dbconfig.json` 保存位置：`app.getPath('userData')` 下（系统用户数据目录）。

## 安装依赖
```powershell
Set-Location "F:\ai\import-database"
npm install
```

## 开发模式 (推荐)
自动同时启动 Vite 与 Electron，窗口加载本地开发服务器：
```powershell
npm run dev
```
如果出现端口占用，可自定义：
```powershell
$env:VITE_PORT=5174; npm run dev
```

## 生产预览
先构建前端再启动 Electron：
```powershell
npm run build:renderer
npm start
```
若 `dist/index.html` 尚未生成，`main.js` 会回退加载根目录 `index.html`，避免 `ERR_FILE_NOT_FOUND`。

## 数据库配置与连接测试
表单字段：Host / Port / User / Password / Database。
操作说明：
- 点击“保存配置”写入 `dbconfig.json`。
- 点击“测试连接”执行 `SELECT 1 AS alive` 验证连接。

### 安全注意事项
- 当前密码为明文存储，生产环境建议：
	- 使用系统凭据管理（Windows Credential Manager / macOS Keychain）。
	- 或对敏感字段进行加密（如使用 `crypto` 对称加密 + 主密钥）。
- 避免在渲染进程直接引入数据库驱动，保持最小攻击面。

## 常见问题
1. `ERR_FILE_NOT_FOUND`：未构建生产文件，请运行 `npm run build:renderer` 或使用 `npm run dev`。
2. “连接失败”错误：检查端口、用户名、密码、数据库是否存在，或 PostgreSQL 是否允许远程连接 (`pg_hba.conf`).
3. 端口被占用：Vite 默认 5173，修改环境变量或释放端口。
4. 权限问题：若无法写入配置，检查用户数据目录权限。

## 后续扩展建议
- 增加多环境配置（保存多个连接条目）。
- 添加加密存储与自动过期策略。
- 增加连接池与执行 SQL 示例。
- 引入日志记录（winston/pino）。

## 快速脚本
```powershell
# 开发
npm run dev
# 构建生产页面
npm run build:renderer
# 生产启动
npm start
```

## 诊断
```powershell
# 查看保存的配置路径（示例输出）
node -e "console.log(require('electron').app.getPath('userData'))"  # 需在运行 Electron 环境中使用
```

欢迎继续提出功能需求。
