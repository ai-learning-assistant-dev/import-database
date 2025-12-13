const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('bridge', {
  getTime: () => ipcRenderer.invoke('get-time'),
  ping: (msg) => ipcRenderer.invoke('ping', msg),
  loadConfig: () => ipcRenderer.invoke('load-db-config'),
  saveConfig: (cfg) => ipcRenderer.invoke('save-db-config', cfg),
  testConnection: (cfg) => ipcRenderer.invoke('test-db-connection', cfg),
  parseExcel: (filePath) => ipcRenderer.invoke('parse-excel', filePath),
  chooseFile: () => ipcRenderer.invoke('choose-file'),
  chooseFolder: () => ipcRenderer.invoke('choose-folder'),
  fetchVideoTitles: (urls) => ipcRenderer.invoke('fetch-video-titles', urls),
  fetchImageBase64: (url) => ipcRenderer.invoke('fetch-image-base64', url),
  queryCourses: (search, limit) => ipcRenderer.invoke('query-courses', { search, limit }),
  upsertCourse: (data) => ipcRenderer.invoke('upsert-course', data),
  importChapters: (courseId, chapters) => ipcRenderer.invoke('import-chapters', { courseId, chapters }),
  importSections: (courseId, rows, folderPath) => ipcRenderer.invoke('import-sections', { courseId, rows, folderPath }),
  clearCourseData: (courseId) => ipcRenderer.invoke('clear-course-data', { courseId }),
  removeCourse: (courseId) => ipcRenderer.invoke('remove-course', { courseId }),
  clearSections: () => ipcRenderer.invoke('clear-sections')
});
