import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import 'antd/dist/reset.css';
import { Layout, Typography, Space, Form, Input, InputNumber, Button, Divider, Table, Tag, message, Card, Tooltip, Modal } from 'antd';
import { Select } from 'antd';

function App() {
  const [time, setTime] = useState('加载中...');
  const [pingResult, setPingResult] = useState('');
  const [cfg, setCfg] = useState({ host:'localhost', port:5432, user:'', password:'', database:'' });
  const [loadStatus, setLoadStatus] = useState('');
  const [saveStatus, setSaveStatus] = useState('');
  const [testStatus, setTestStatus] = useState('');
  const [excelPath, setExcelPath] = useState('');
  const [excelStatus, setExcelStatus] = useState('');
  const [videoList, setVideoList] = useState([]);
  const [fetchingTitles, setFetchingTitles] = useState(false);
  const [sectionsFolder, setSectionsFolder] = useState('');
  const [importDetails, setImportDetails] = useState([]);
  const [importMissing, setImportMissing] = useState([]);
  const [importStats, setImportStats] = useState(null);
  const [detailsModalOpen, setDetailsModalOpen] = useState(false);
  // Courses state
  const [courseSearch, setCourseSearch] = useState('');
  const [courses, setCourses] = useState([]);
  const [loadingCourses, setLoadingCourses] = useState(false);
  const [selectedCourseId, setSelectedCourseId] = useState(null);
  const [selectedCourse, setSelectedCourse] = useState(null);
  const [courseForm, setCourseForm] = useState({ name:'', icon_url:'', description:'', default_ai_persona_id:'' });
  const [upserting, setUpserting] = useState(false);
  const [isUpdateMode, setIsUpdateMode] = useState(false);
  const [initialCoursesLoaded, setInitialCoursesLoaded] = useState(false);
  const [courseModalOpen, setCourseModalOpen] = useState(false);

  const refreshTime = async () => {
    try {
      const t = await window.bridge.getTime();
      setTime(t);
    } catch (e) {
      setTime('获取失败');
    }
  };

  const sendPing = async () => {
    const res = await window.bridge.ping('hello');
    setPingResult(res);
  };

  useEffect(() => {
    refreshTime();
    (async () => {
      const existing = await window.bridge.loadConfig();
      if (existing && !existing.error) {
        setCfg(c => ({...c, ...existing}));
        setLoadStatus('已加载保存配置');
      } else if (existing && existing.error) {
        setLoadStatus('加载失败: ' + existing.error);
      } else {
        setLoadStatus('无已保存配置');
      }
      // 初次自动查询课程列表
      queryCourses('');
    })();
  }, []);

  const updateField = (field, value) => {
    setCfg(c => ({ ...c, [field]: value }));
  };

  const saveConfig = async () => {
    setSaveStatus('保存中...');
    const res = await window.bridge.saveConfig(cfg);
    if (res.ok) { setSaveStatus('保存成功'); message.success('保存成功'); } else { setSaveStatus('保存失败: ' + res.error); message.error(res.error); }
  };

  const testConnection = async () => {
    setTestStatus('测试连接中...');
    const res = await window.bridge.testConnection(cfg);
    if (res.ok) { setTestStatus('连接成功'); message.success('数据库连接成功'); } else { setTestStatus('连接失败: ' + res.error); message.error(res.error); }
  };

  // Excel 解析
  const chooseFile = async () => {
    const path = await window.bridge.chooseFile();
    if (path) {
      setExcelPath(path);
      message.info('已选择: ' + path);
    }
  };
  const handleExcelParse = async () => {
    if (!excelPath) { setExcelStatus('请先选择文件'); message.warning('请先选择文件'); return; }
    setExcelStatus('解析中...');
    const res = await window.bridge.parseExcel(excelPath);
    if (!res.ok) {
      setExcelStatus('解析失败: ' + res.error);
      setVideoList([]);
      message.error(res.error);
      return;
    }
    setExcelStatus(`解析成功，共 ${res.list.length} 条，正在抓取标题...`);
    setVideoList(res.list);
    if (!res.list.length) {
      message.success('解析成功，但没有可用数据');
      return;
    }
    try {
      setFetchingTitles(true);
      const urls = res.list.map(v => v.url).filter(Boolean);
      if (!urls.length) {
        message.info('解析成功，但未找到可抓取的 URL');
        setExcelStatus(`解析成功，共 ${res.list.length} 条`);
        return;
      }
      const fetchRes = await window.bridge.fetchVideoTitles(urls);
      if (Array.isArray(fetchRes)) {
        const map = new Map(fetchRes.map(r => [r.url, r]));
        setVideoList(list => list.map(item => {
          const r = map.get(item.url);
          if (!r) return item;
          return {
            ...item,
            title: r.ok ? (r.title || item.title) : item.title,
            imageSrc: r.imageSrc || item.imageSrc,
            error: r.ok ? null : r.error
          };
        }));
        const okCount = fetchRes.filter(r => r.ok).length;
        setExcelStatus(`解析成功，共 ${res.list.length} 条；标题抓取成功 ${okCount}/${fetchRes.length}`);
        message.success(`解析 + 抓取完成：成功 ${okCount}/${fetchRes.length}`);
      } else {
        setExcelStatus(`解析成功，共 ${res.list.length} 条；标题抓取返回格式异常`);
        message.error('抓取标题返回数据格式异常');
      }
    } catch (e) {
      setExcelStatus(`解析成功，共 ${res.list.length} 条；抓取标题失败：${e.message}`);
      message.error('抓取标题失败: ' + e.message);
    } finally {
      setFetchingTitles(false);
    }
  };

  const videoColumns = [
    { title: '序号', dataIndex: 'index', width: 60, render: (_v,_r,i) => i + 1 },
    { title: '章标题', dataIndex: '章标题', render: (v,record) => v || <Tag color="orange">(空)</Tag> }, 
    { title: '章顺序', dataIndex: '章顺序', render: (v,record) => v || <Tag color="orange">(空)</Tag> },
    { title: '节标题', dataIndex: '节标题', render: (v,record) => v || <Tag color="orange">(空)</Tag> },
    { title: '节顺序', dataIndex: '节顺序', render: (v,record) => v || <Tag color="orange">(空)</Tag> },
    { title: '机械标题', dataIndex: '机械标题', render: (v,record) => v || <Tag color="orange">(空)</Tag> }, 
    { title: '节图片', dataIndex: 'imageSrc', render: (v,record) => v ? <Typography.Text copyable={{text:v}} ellipsis={{tooltip:v}} style={{maxWidth:260,display:'inline-block'}}>{v}</Typography.Text> : '-' },
    { title: '网页标题', dataIndex: 'title', width: 240, render: (_v, record) => {
        if (record.error) {
          return <Tooltip title={record.error}><Tag color="red">抓取失败</Tag></Tooltip>;
        }
        return record.title || <Tag color="gold">待抓取</Tag>;
      }
    },
    { title: '节视频', dataIndex: 'URL', render: (_v,record) => <Typography.Text copyable={{text:record.url}} ellipsis={{tooltip:record.url}} style={{maxWidth:300,display:'inline-block'}}>{record.url}</Typography.Text> }
  ];

  const fetchTitles = async () => {
    if (!videoList.length) { message.warning('没有可抓取的 URL'); return; }
    setFetchingTitles(true);
    try {
      const urls = videoList.map(v => v.url).filter(Boolean);
      const res = await window.bridge.fetchVideoTitles(urls);
      if (Array.isArray(res)) {
        const map = new Map(res.map(r => [r.url, r]));
        setVideoList(list => list.map(item => {
          const r = map.get(item.url);
          if (!r) return item;
          return { ...item, title: r.ok ? (r.title || item.title) : item.title, imageSrc: r.imageSrc || item.imageSrc, error: r.ok ? null : r.error };
        }));
        const okCount = res.filter(r => r.ok).length;
        message.success(`抓取完成 成功 ${okCount}/${res.length}`);
      } else {
        message.error('返回数据格式异常');
      }
    } catch (e) {
      message.error('抓取失败: ' + e.message);
    } finally {
      setFetchingTitles(false);
    }
  };

  // 查询 courses
  const queryCourses = async (manualTerm) => {
    const term = manualTerm !== undefined ? manualTerm : courseSearch;
    setLoadingCourses(true);
    try {
      const res = await window.bridge.queryCourses(term, 50);
      if (res.ok) {
        setCourses(res.list || []);
        if ((res.list || []).length === 0) message.info('未找到课程');
        if (!initialCoursesLoaded) {
          if (res.list && res.list.length > 0) {
            const first = res.list[0];
            setSelectedCourseId(first.course_id);
            setSelectedCourse(first);
            setIsUpdateMode(true);
            setCourseForm({
              name: first.name || '',
              icon_url: first.icon_url || '',
              description: first.description || '',
              default_ai_persona_id: first.default_ai_persona_id ?? ''
            });
          } else {
            // 无数据进入新增模式
            setIsUpdateMode(false);
            setCourseForm({ name:'', icon_url:'', description:'', default_ai_persona_id:'' });
            setSelectedCourseId(null);
            setSelectedCourse(null);
          }
          setInitialCoursesLoaded(true);
        }
      } else {
        message.error('课程查询失败: ' + res.error);
        setCourses([]);
      }
    } catch (e) {
      message.error('查询异常: ' + e.message);
      setCourses([]);
    } finally {
      setLoadingCourses(false);
    }
  };

  const handleSelectCourse = async (courseId) => {
    setSelectedCourseId(courseId);
    const c = courses.find(c => String(c.course_id) === String(courseId));
    setSelectedCourse(c || null);
    if (c) {
      setCourseForm({
        name: c.name || '',
        icon_url: c.icon_url || '',
        description: c.description || '',
        default_ai_persona_id: c.default_ai_persona_id ?? ''
      });
      setIsUpdateMode(true);
    }
  };

  const courseOptions = courses.map(c => ({
    label: c.name || ('课程 ' + c.course_id),
    value: c.course_id
  }));

  const onCourseFormChange = (field, value) => {
    setCourseForm(f => ({ ...f, [field]: value }));
  };

  const submitCourse = async () => {
    if (!courseForm.name.trim()) { message.warning('请输入课程名称'); return; }
    if (!courseForm.icon_url.trim()) { message.warning('请输入课程图片URL'); return; }
    if (!isUpdateMode) {
      const exists = courses.some(c => c.name === courseForm.name.trim());
      if (exists) { message.error('名称已存在，请通过下拉选择进行修改'); return; }
    }
    setUpserting(true);
    try {
      const res = await window.bridge.upsertCourse({
        name: courseForm.name.trim(),
        icon_url: courseForm.icon_url.trim(),
        description: courseForm.description.trim(),
        default_ai_persona_id: courseForm.default_ai_persona_id
      });
      if (res.ok) {
        message.success(res.action === 'update' ? '课程已更新' : '课程已新增');
        // 刷新下拉：保持当前搜索关键字，不强制只显示新名称
        await queryCourses(courseSearch.trim());
        if (res.course) {
          setSelectedCourseId(res.course.course_id);
          setSelectedCourse(res.course);
        }
        setIsUpdateMode(res.action === 'update');
        setCourseModalOpen(false);
      } else {
        message.error('保存失败: ' + res.error);
      }
    } catch (e) {
      message.error('保存异常: ' + e.message);
    } finally {
      setUpserting(false);
    }
  };

  const importSectionsConfirm = () => {
    Modal.confirm({
      title: '确认导入视频和习题？',
      content: (
        <div>
          <Typography.Title level={4} style={{ marginTop: 0 }}>
            课程：{selectedCourse?.name || '未命名课程'}
          </Typography.Title>
          <Typography.Paragraph>
            将根据当前 Excel 列表和素材文件夹导入章节与对应的节/习题，确认继续？
          </Typography.Paragraph>
        </div>
      ),
      okText: '开始导入',
      cancelText: '取消',
      async onOk() {
        // 构建章节列表
        const map = new Map();
        const chapters = [];
        for (const row of videoList) {
          const rawTitle = (row['章标题'] || '').trim();
          if (!rawTitle) continue;
          if (map.has(rawTitle)) continue; // 去重同一 Excel 中重复
          map.set(rawTitle, true);
          const orderRaw = row['章顺序'];
          let orderNum = parseInt(orderRaw, 10);
          if (!Number.isFinite(orderNum)) orderNum = chapters.length + 1;
          chapters.push({ title: rawTitle, chapter_order: orderNum });
        }
        if (!chapters.length) {
          message.warning('没有可导入的章节');
          return;
        }

        const hideChapters = message.loading('正在导入章节...', 0);
        try {
          const resChapters = await window.bridge.importChapters(selectedCourseId, chapters);
          hideChapters();
          if (resChapters.ok) {
            if (resChapters.inserted > 0) {
              message.success(`章节导入完成，新增 ${resChapters.inserted} 条，跳过 ${resChapters.skipped} 条，继续导入节...`);
            } else {
              message.info('无新增章节，可能都已存在或为空，继续导入节...');
            }
          } else {
            message.error('章节导入失败: ' + resChapters.error);
            return; // 章节失败就不继续导入节
          }
        } catch (e) {
          hideChapters();
          message.error('章节导入异常: ' + e.message);
          return;
        }

        const hideSections = message.loading('正在导入节 (sections)...', 0);
        try {
          const resSections = await window.bridge.importSections(selectedCourseId, videoList, sectionsFolder);
          hideSections();
          if (resSections.ok) {
            setImportDetails(Array.isArray(resSections.details) ? resSections.details : []);
            setImportMissing(Array.isArray(resSections.missingFiles) ? resSections.missingFiles : []);
            setImportStats({ inserted: resSections.inserted, skipped: resSections.skipped });
            const missCount = Array.isArray(resSections.missingFiles) ? resSections.missingFiles.length : 0;
            message.success(`节导入完成 新增 ${resSections.inserted} 条, 跳过 ${resSections.skipped} 条, 文件缺失或错误 ${missCount} 条`);
            if (missCount) {
              console.warn('missingFiles', resSections.missingFiles);
            }
          } else {
            setImportDetails([]);
            setImportMissing([]);
            setImportStats(null);
            message.error('节导入失败: ' + resSections.error);
          }
        } catch (e) {
          hideSections();
          setImportDetails([]);
          setImportMissing([]);
          setImportStats(null);
          message.error('节导入异常: ' + e.message);
        }
      }
    });

  }

  const clearSectionsConfirm = () => {
    Modal.confirm({
      title: '确认重新导入该课程？',
      content: (
        <div>
          <Typography.Title level={4} style={{ marginTop: 0 }}>
            将清除当前课程的所有章节、节和习题数据
          </Typography.Title>
          <Typography.Paragraph>
            课程：{selectedCourse.name}（ID: {selectedCourse.course_id}）
          </Typography.Paragraph>
          <Typography.Paragraph type="danger">
            操作会删除该课程下的 chapters、sections、leading_question、exercises、exercise_options 中的相关记录，且不可恢复，确认继续？
          </Typography.Paragraph>
        </div>
      ),
      okText: '确认清除',
      cancelText: '取消',
      okButtonProps: { danger: true },
      onOk: async () => {
        try {
          const res = await window.bridge.clearCourseData(selectedCourseId);
          if (res.ok) {
            message.success(`已清除课程【${selectedCourse.name}】下的所有章节和习题数据`);
          } else {
            message.error('清除失败: ' + res.error);
          }
        } catch (e) {
          message.error('清除异常: ' + e.message);
        }
      }
    });
  }
  

  return (
    <>
    <Layout style={{minHeight:'100vh'}}>
      <Layout.Header style={{background:'#141414'}}>
        <Typography.Title level={4} style={{color:'#fff',margin:0}}>管理面板</Typography.Title>
      </Layout.Header>
      <Layout.Content style={{padding:24}}>
        <Space direction="vertical" size={24} style={{width:'100%'}}>
          {/* 数据库配置放在最上面 */}
          <Card title="数据库配置" size="small" extra={<Typography.Text type="secondary">保存位置: userData</Typography.Text>}>
            <Form layout="inline" style={{rowGap:12}}>
              <Form.Item label="Host">
                <Input value={cfg.host} onChange={e=>updateField('host', e.target.value)} style={{width:160}}/>
              </Form.Item>
              <Form.Item label="Port">
                <InputNumber value={cfg.port} onChange={v=>updateField('port', v)} style={{width:100}}/>
              </Form.Item>
              <Form.Item label="User">
                <Input value={cfg.user} onChange={e=>updateField('user', e.target.value)} style={{width:140}}/>
              </Form.Item>
              <Form.Item label="Password">
                <Input.Password value={cfg.password} onChange={e=>updateField('password', e.target.value)} style={{width:180}}/>
              </Form.Item>
              <Form.Item label="Database">
                <Input value={cfg.database} onChange={e=>updateField('database', e.target.value)} style={{width:180}}/>
              </Form.Item>
              <Form.Item>
                <Space>
                  <Button type="primary" onClick={saveConfig}>保存配置</Button>
                  <Button onClick={testConnection}>测试连接</Button>
                </Space>
              </Form.Item>
            </Form>
            <Divider style={{margin:'12px 0'}}/>
            <Space>
              <Tag color="blue">{loadStatus}</Tag>
              <Tag color="green">{saveStatus}</Tag>
              <Tag color="purple">{testStatus}</Tag>
            </Space>
          </Card>
          <Card title="选择要导入的课程" size="small" extra={<Typography.Text type="secondary">来自数据库 courses 表</Typography.Text>}>
            <Space wrap align="start">
              <Input.Search
                allowClear
                placeholder="输入课程名称关键字"
                value={courseSearch}
                onChange={e=>setCourseSearch(e.target.value)}
                onSearch={(v)=>{ setCourseSearch(v); queryCourses(v); }}
                style={{width:240}}
                enterButton="查询"
                loading={loadingCourses}
              />
              <Select
                showSearch
                placeholder="选择课程"
                options={courseOptions}
                value={selectedCourseId}
                onChange={handleSelectCourse}
                style={{minWidth:260}}
                filterOption={(input, option)=> (option?.label ?? '').toLowerCase().includes(input.toLowerCase())}
                loading={loadingCourses}
                notFoundContent={loadingCourses ? '加载中...' : '无数据'}
              />
              <Button onClick={()=>queryCourses()} disabled={loadingCourses}>刷新</Button>
              <Button type="dashed" onClick={()=>{ setIsUpdateMode(false); setSelectedCourseId(null); setSelectedCourse(null); setCourseForm({ name:'', icon_url:'', description:'', default_ai_persona_id:'' }); setCourseModalOpen(true); }}>新增课程</Button>
              <Button danger onClick={async ()=>{
                if (!selectedCourseId || !selectedCourse) { message.warning('请先选择课程'); return; }
                Modal.confirm({
                  title: '确认删除该课程？',
                  content: (
                    <div>
                      删除后会导致学生的学习记录丢失
                    </div>
                  ),
                  okText: '确认删除',
                  cancelText: '取消',
                  okButtonProps: { danger: true },
                  onOk: async () => {
                    try {
                      const res = await window.bridge.removeCourse(selectedCourseId);
                      if (res.ok) {
                        message.success(`已删除课程【${selectedCourse.name}】`);
                        handleSelectCourse(null);
                        queryCourses();
                      } else {
                        message.error('删除失败: ' + res.error);
                      }
                    } catch (e) {
                      message.error('删除异常: ' + e.message);
                    }
                  }
                });
              }}>删除课程</Button>
            </Space>
            {selectedCourse && (
              <Divider style={{margin:'12px 0'}} />
            )}
            {selectedCourse && (
              <Space align="start" style={{display:'flex'}}>
                <div style={{flex:1}}>
                  <Typography.Title level={5} style={{marginTop:0}}>{selectedCourse.name} <Tag>#{selectedCourse.course_id}</Tag></Typography.Title>
                  <Typography.Paragraph ellipsis={{rows:3, expandable:true, symbol:'展开'}}> {selectedCourse.description || '无描述'} </Typography.Paragraph>
                  <Space size="small" wrap>
                    <Tag color="blue">default_ai_persona_id: {selectedCourse.default_ai_persona_id ?? 'null'}</Tag>
                    {selectedCourse.icon_url && (
                      <Button
                        size="small"
                        onClick={() => navigator.clipboard && navigator.clipboard.writeText(selectedCourse.icon_url)}
                      >
                        复制封面URL
                      </Button>
                    )}
                    <Button size="small" type="primary" onClick={()=>{
                      // 将当前选中课程同步到表单并打开编辑模态框
                      setIsUpdateMode(true);
                      setCourseForm({
                        name: selectedCourse.name || '',
                        icon_url: selectedCourse.icon_url || '',
                        description: selectedCourse.description || '',
                        default_ai_persona_id: selectedCourse.default_ai_persona_id ?? ''
                      });
                      setCourseModalOpen(true);
                    }}>修改课程</Button>
                  </Space>
                </div>
              </Space>
            )}
          </Card>
          <Card title="Excel 视频列表" size="small" extra={<Typography.Text type="secondary">支持 xlsx/xls</Typography.Text>}>
            <Space wrap>
              <Button onClick={chooseFile}>选择文件</Button>
              <Input style={{width:400}} placeholder="已选择文件路径" value={excelPath} readOnly/>
              <Button type="primary" onClick={handleExcelParse} loading={fetchingTitles}>解析文件并抓取标题</Button>
              <Button onClick={async()=>{ const p = await window.bridge.chooseFolder(); if (p) { setSectionsFolder(p); message.success('已选择目录'); } }}>
                选择素材文件夹
              </Button>
              <Input style={{width:360}} placeholder="素材文件夹路径" value={sectionsFolder} readOnly />
              <Button
                type="primary"
                disabled={fetchingTitles || !videoList.length || !selectedCourseId || !sectionsFolder}
                onClick={async()=>{
                  if (!selectedCourseId) { message.warning('请先选择课程'); return; }
                  if (!videoList.length) { message.warning('请先解析 Excel'); return; }
                  if (!sectionsFolder) { message.warning('请选择素材文件夹'); return; }
                  if ((excelPath.indexOf(selectedCourse?.name) < 0 || sectionsFolder.indexOf(selectedCourse?.name) < 0)) {
                    Modal.confirm({
                      title: ' ！！！你可能选错了课程，请仔细检查！！！',
                      okText: '返回',
                      cancelText: '我检查无误',
                      onCancel: importSectionsConfirm
                    })
                  } else {
                    importSectionsConfirm();
                  }
                }}
              >导入章节和习题</Button>
              <Button danger onClick={async ()=>{
                if (!selectedCourseId || !selectedCourse) { message.warning('请先选择课程'); return; }
                if ((excelPath.indexOf(selectedCourse?.name) < 0 || sectionsFolder.indexOf(selectedCourse?.name) < 0)) {
                  Modal.confirm({
                    title: ' ！！！你可能选错了课程，请仔细检查！！！',
                    okText: '返回',
                    cancelText: '我检查无误',
                    onCancel: clearSectionsConfirm
                  })
                } else {
                  clearSectionsConfirm();
                }
                
              }}>重新导入（清除当前课程数据）</Button>
              <Button disabled={!(importDetails.length || importMissing.length)} onClick={()=>setDetailsModalOpen(true)}>查看导入详情</Button>
              <Typography.Text type="secondary">{excelStatus}</Typography.Text>
            </Space>
            <Divider style={{margin:'12px 0'}}/>
            <Table
              size="small"
              dataSource={videoList.map((v,i)=>({...v,key:i}))}
              columns={videoColumns}
              pagination={{pageSize:10}}
              locale={{emptyText:'暂无数据'}}
            />
          </Card>
          <Typography.Text type="secondary" style={{marginTop:8}}>提示：密码暂为明文，建议后续加密存储。</Typography.Text>
        </Space>
      </Layout.Content>
    </Layout>
    {/* 课程新增 / 修改模态框 */}
    <Modal
      title={isUpdateMode ? '修改课程' : '新增课程'}
      open={courseModalOpen}
      onCancel={()=>{ setCourseModalOpen(false); }}
      footer={null}
      destroyOnClose
    >
      <Form layout="vertical" onFinish={submitCourse}>
        <Form.Item label="课程名称" required validateStatus={!courseForm.name.trim() ? 'error' : ''} help={!courseForm.name.trim() ? '必填' : ''}>
          <Input value={courseForm.name} onChange={e=>onCourseFormChange('name', e.target.value)} placeholder="唯一名称" allowClear />
        </Form.Item>
        <Form.Item label="图标URL" required validateStatus={!courseForm.icon_url.trim() ? 'error' : ''} help={!courseForm.icon_url.trim() ? '必填' : ''}>
          <Input value={courseForm.icon_url} onChange={e=>onCourseFormChange('icon_url', e.target.value)} placeholder="https://..." allowClear />
        </Form.Item>
        <Form.Item label="描述">
          <Input.TextArea rows={3} value={courseForm.description} onChange={e=>onCourseFormChange('description', e.target.value)} placeholder="可选描述" />
        </Form.Item>
        <Form.Item label="默认 AI Persona ID">
          <InputNumber style={{width:'100%'}} value={courseForm.default_ai_persona_id} onChange={v=>onCourseFormChange('default_ai_persona_id', v)} placeholder="可选数值" />
        </Form.Item>
        <Form.Item>
          <Space>
            <Button type="primary" htmlType="submit" loading={upserting}>{isUpdateMode ? '更新课程' : '新增课程'}</Button>
            <Button onClick={()=>{ setCourseForm({ name:'', icon_url:'', description:'', default_ai_persona_id:'' }); setIsUpdateMode(false); setSelectedCourseId(null); setSelectedCourse(null); }}>重置</Button>
          </Space>
        </Form.Item>
      </Form>
    </Modal>
    <Modal
      title="节导入详情日志"
      open={detailsModalOpen}
      onCancel={()=>setDetailsModalOpen(false)}
      footer={null}
      width={720}
    >
      <Space direction="vertical" size="large" style={{width:'100%'}}>
        {importStats && (
          <Typography.Text type="secondary">本次导入：新增 {importStats.inserted} 条，跳过 {importStats.skipped} 条。</Typography.Text>
        )}
        <div>
          <Typography.Title level={5} style={{marginTop:0}}>缺失或无法解析的文件</Typography.Title>
          <Table
            size="small"
            dataSource={importMissing.map((m,i)=>({ ...m, key:i }))}
            columns={[
              { title:'节标题', dataIndex:'sectionTitle', width:200 },
              { title:'原因', dataIndex:'reason', render:(v)=>v ? <Tag color="red">{v}</Tag> : '-' }
            ]}
            pagination={{pageSize:10}}
            locale={{emptyText:'无缺失或错误文件'}}
          />
        </div>
      </Space>
    </Modal>
    </>
  );
}

// 旧的内联样式已移除，使用 AntD 组件替代

createRoot(document.getElementById('root')).render(<App />);
