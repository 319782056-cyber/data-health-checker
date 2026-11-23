import { useMemo, useRef, useState } from 'react'
import type { Dispatch, SetStateAction, ReactNode } from 'react'
import { HashRouter, NavLink, Route, Routes, useLocation } from 'react-router-dom'
import * as XLSX from 'xlsx'
import {
  Activity, ArrowDown, ArrowUp, BarChart3, Bot, Check, ChevronDown,
  Combine, Database, Download, FilePlus2, FileSpreadsheet, History, LayoutDashboard,
  Menu, MoreHorizontal, Search, Settings2, ShieldCheck, Sparkles, Table2,
  WandSparkles, Workflow, X
} from 'lucide-react'
import './v2.css'

type Cell = string | number | boolean | Date | null
type Dataset = { id: string; fileName: string; sheetName: string; headers: string[]; rows: Cell[][] }
type Mode = 'simple' | 'professional'
type Step = { id: string; label: string; detail: string }
const uid = () => Math.random().toString(36).slice(2, 10)
const cellText = (v: Cell) => v instanceof Date ? v.toISOString().slice(0, 10) : String(v ?? '')
const blank = (v: Cell) => v == null || String(v).trim() === ''
type FieldQuality = { name: string; missing: number; filled: number; unique: number; type: string; completeness: number }
function inspectQuality(dataset: Dataset) {
  const totalRows = dataset.rows.length
  const duplicateRows = totalRows - new Set(dataset.rows.map(row => JSON.stringify(row.map(v => v instanceof Date ? v.toISOString() : v)))).size
  const fields: FieldQuality[] = dataset.headers.map((name, index) => {
    const values = dataset.rows.map(row => row[index])
    const filledValues = values.filter(v => !blank(v))
    const missing = totalRows - filledValues.length
    const unique = new Set(filledValues.map(v => cellText(v).trim())).size
    const kinds = new Set(filledValues.map(v => {
      if (v instanceof Date) return '日期'
      if (typeof v === 'boolean') return '布尔'
      if (typeof v === 'number') return '数值'
      const value = String(v).trim()
      if (value && !Number.isNaN(Number(value))) return '数值'
      if (/^\d{4}[-/]\d{1,2}([-/]\d{1,2})?$/.test(value)) return '日期'
      return '文本'
    }))
    const type = kinds.size === 0 ? '空字段' : kinds.size === 1 ? [...kinds][0] : '混合类型'
    return { name, missing, filled: filledValues.length, unique, type, completeness: totalRows ? Math.round(filledValues.length / totalRows * 100) : 100 }
  })
  const totalCells = totalRows * dataset.headers.length
  const missingCells = fields.reduce((sum, field) => sum + field.missing, 0)
  const completeness = totalCells ? Math.round((totalCells - missingCells) / totalCells * 100) : 100
  return { totalRows, duplicateRows, missingCells, totalCells, completeness, fields }
}
function exportQualityReport(dataset: Dataset) {
  const quality = inspectQuality(dataset)
  const wb = XLSX.utils.book_new()
  const rows: (string | number)[][] = [
    ['数据质量报告', dataset.fileName + ' / ' + dataset.sheetName],
    ['总行数', quality.totalRows], ['字段数', dataset.headers.length],
    ['完整度', quality.completeness + '%'], ['空值单元格', quality.missingCells],
    ['完全重复行', quality.duplicateRows], [],
    ['字段', '识别类型', '非空数量', '空值数量', '唯一值数量', '完整度']
  ]
  quality.fields.forEach(field => rows.push([field.name, field.type, field.filled, field.missing, field.unique, field.completeness + '%']))
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '质量报告')
  XLSX.writeFile(wb, dataset.fileName.replace(/\.[^.]+$/, '') + '_数据质量报告.xlsx')
}

const nav = [
  { to: '/overview', label: '数据总览', icon: LayoutDashboard },
  { to: '/data', label: '数据浏览', icon: Table2 },
  { to: '/transform', label: '智能处理', icon: WandSparkles },
  { to: '/merge', label: '多表合并', icon: Combine },
  { to: '/analysis', label: '数据分析', icon: BarChart3 },
  { to: '/workflows', label: '自动化流程', icon: Workflow },
  { to: '/activity', label: '操作记录', icon: History },
]
const titleMap: Record<string, string> = Object.fromEntries(nav.map(n => [n.to, n.label]))

function AppShell() {
  const [datasets, setDatasets] = useState<Dataset[]>([])
  const [activeId, setActiveId] = useState('')
  const [mode, setMode] = useState<Mode>('simple')
  const [steps, setSteps] = useState<Step[]>([])
  const [history, setHistory] = useState<string[]>([])
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)
  const current = datasets.find(d => d.id === activeId) ?? datasets[0]
  const location = useLocation()
  const pageTitle = titleMap[location.pathname] ?? '数据总览'

  async function importFiles(files: FileList | null) {
    if (!files?.length) return
    setBusy(true); setMessage('')
    try {
      const added: Dataset[] = []
      for (const file of Array.from(files)) {
        const ext = file.name.split('.').pop()?.toLowerCase()
        if (!['xlsx', 'xls', 'csv'].includes(ext ?? '')) continue
        const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true })
        workbook.SheetNames.forEach(sheetName => {
          const raw = XLSX.utils.sheet_to_json<Cell[]>(workbook.Sheets[sheetName], { header: 1, defval: null, raw: true })
          if (!raw.length) return
          const width = Math.max(...raw.map(row => row.length))
          const headers = Array.from({ length: width }, (_, i) => cellText(raw[0][i]).trim() || '字段 ' + (i + 1))
          added.push({ id: uid(), fileName: file.name, sheetName, headers, rows: raw.slice(1).map(row => Array.from({ length: width }, (_, i) => row[i] ?? null)) })
        })
      }
      if (!added.length) throw new Error('没有读取到有效工作表，请检查文件格式。')
      setDatasets(prev => [...prev, ...added]); setActiveId(added[0].id)
      setMessage('已导入 ' + added.length + ' 个工作表。')
    } catch (error) { setMessage(error instanceof Error ? error.message : '导入失败') }
    finally { setBusy(false); if (fileRef.current) fileRef.current.value = '' }
  }
  function record(label: string) { setHistory(prev => [label, ...prev]); setMessage(label + '：已记录。') }
  function exportCurrent() {
    if (!current) return
    const rows = [current.headers, ...current.rows].map(row => row.map(v => v instanceof Date ? v.toISOString() : v))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), current.sheetName.slice(0, 31) || '数据')
    XLSX.writeFile(wb, current.fileName.replace(/\.[^.]+$/, '') + '_导出.xlsx')
  }
  const context = { datasets, setDatasets, activeId, setActiveId, current, mode, setMode, steps, setSteps, history, record, importFiles, openImport: () => fileRef.current?.click(), exportCurrent }
  return <div className="v2-shell">
    <aside className={'v2-sidebar ' + (sidebarOpen ? 'open' : '')}>
      <div className="v2-brand"><div className="v2-brand-icon"><Activity size={21}/></div><div><b>Data <span>Workbench</span></b><small>INTELLIGENT DATA STUDIO</small></div><button className="v2-mobile-close" onClick={() => setSidebarOpen(false)}><X size={18}/></button></div>
      <div className="v2-nav-label">工作空间</div>
      <nav>{nav.map(item => { const Icon = item.icon; return <NavLink key={item.to} to={item.to} onClick={() => setSidebarOpen(false)} className={({ isActive }) => 'v2-nav-item ' + (isActive ? 'active' : '')}><Icon size={17}/><span>{item.label}</span>{item.to === '/activity' && history.length > 0 && <i>{history.length}</i>}</NavLink> })}</nav>
      <div className="v2-sidebar-bottom">
        <div className="v2-privacy"><ShieldCheck size={17}/><div><b>本地优先模式</b><small>云端 AI 需单独授权</small></div><Check size={14}/></div>
        <NavLink to="/settings" className="v2-settings-link"><Settings2 size={16}/> 设置与隐私</NavLink>
        <div className="v2-version">DATA WORKBENCH · 2.0 FOUNDATION</div>
      </div>
    </aside>
    {sidebarOpen && <button aria-label="关闭导航" className="v2-scrim" onClick={() => setSidebarOpen(false)}/>}
    <main className="v2-main">
      <header className="v2-topbar"><div className="v2-top-left"><button className="v2-menu" onClick={() => setSidebarOpen(true)}><Menu size={19}/></button><span>工作空间</span><b>/</b><strong>{pageTitle}</strong></div><div className="v2-top-actions">{busy && <span className="v2-importing">正在读取文件…</span>}<div className="v2-mode-switch"><button className={mode === 'simple' ? 'selected' : ''} onClick={() => setMode('simple')}>简单模式</button><button className={mode === 'professional' ? 'selected' : ''} onClick={() => setMode('professional')}>专业模式</button></div><button className="v2-ai-button" onClick={() => setMessage('AI 助手将在云端授权与计划校验模块完成后启用。当前数据不会发送到云端。')}><Sparkles size={15}/> AI 助手</button></div></header>
      <div className="v2-content">
        <Routes>
          <Route path="/" element={<Overview {...context}/>}/>
          <Route path="/overview" element={<Overview {...context}/>}/>
          <Route path="/data" element={<DataBrowser {...context}/>}/>
          <Route path="/transform" element={<Transform {...context}/>}/>
          <Route path="/merge" element={<ModulePage icon={<Combine/>} title="多表合并" description="通过追加或关联，将多个数据来源组合成一个可追溯的数据集。" bullets={['纵向追加：按字段名称对齐多张表','横向关联：Left、Inner、Full、Right Join','字段匹配建议、未匹配记录与行数变化检查']} status="数据引擎接入后开放"/>}/>
          <Route path="/analysis" element={<ModulePage icon={<BarChart3/>} title="数据分析" description="使用分组、聚合和数据透视，把数据转化为可用的分析结果。" bullets={['拖放字段配置行、列和值','求和、计数、平均值、最大值和最小值','透视结果可继续导出或作为新数据集']} status="分析模块开发中"/>}/>
          <Route path="/workflows" element={<ModulePage icon={<Workflow/>} title="自动化流程" description="把数据处理步骤保存为模板，并在新数据到来时重复执行。" bullets={['可编辑、可重放的有序处理步骤','记录输入输出版本与执行结果','检测字段变化并在不安全时暂停']} status="工作流引擎开发中"/>}/>
          <Route path="/activity" element={<ActivityPage history={history}/>}/>
          <Route path="/settings" element={<SettingsPage mode={mode} setMode={setMode}/>}/>
        </Routes>
      </div>
      {message && <div className="v2-toast" role="status">{message}<button onClick={() => setMessage('')}><X size={14}/></button></div>}
      <input ref={fileRef} hidden multiple type="file" accept=".xlsx,.xls,.csv" onChange={e => void importFiles(e.target.files)}/>
    </main>
  </div>
}

type Context = {
  datasets: Dataset[]; setDatasets: Dispatch<SetStateAction<Dataset[]>>;
  activeId: string; setActiveId: (id: string) => void; current?: Dataset;
  mode: Mode; setMode: (mode: Mode) => void; steps: Step[]; setSteps: Dispatch<SetStateAction<Step[]>>;
  history: string[]; record: (label: string) => void; importFiles: (files: FileList | null) => Promise<void>;
  openImport: () => void; exportCurrent: () => void;
}
function PageHeading({ eyebrow, title, description, action }: { eyebrow: string; title: string; description: string; action?: ReactNode }) {
  return <div className="v2-page-heading"><div><div className="v2-eyebrow"><Sparkles size={13}/>{eyebrow}</div><h1>{title}</h1><p>{description}</p></div>{action}</div>
}
function Overview(p: Context) {
  const totalRows = p.datasets.reduce((sum, d) => sum + d.rows.length, 0)
  return <><PageHeading eyebrow="DATA WORKSPACE" title="让数据工作更简单。" description="从文件导入开始，逐步完成识别、清洗、合并与分析。" action={<button className="v2-primary" onClick={p.openImport}><FilePlus2 size={16}/> 导入数据</button>}/>
    <div className="v2-stat-grid"><Stat label="数据集 / 工作表" value={p.datasets.length.toString()} icon={<Database/>}/><Stat label="数据记录" value={totalRows.toLocaleString()} icon={<Table2/>}/><Stat label="处理步骤" value={p.steps.length.toString()} icon={<Workflow/>}/><Stat label="操作记录" value={p.history.length.toString()} icon={<History/>}/></div>
    <section className="v2-card v2-start-card"><div className="v2-start-icon"><FileSpreadsheet size={25}/></div><div className="v2-start-copy"><h2>{p.datasets.length ? '继续你的数据工作' : '创建第一个数据工作区'}</h2><p>{p.datasets.length ? '选择一个数据集继续浏览或处理。' : '支持 Excel 和 CSV 文件。文件默认只在当前浏览器中解析，不会自动上传。'}</p></div><button className="v2-secondary" onClick={p.openImport}><FilePlus2 size={15}/> 选择文件</button></section>
    <section className="v2-card"><div className="v2-section-head"><div><h2>最近数据集</h2><p>当前浏览器会话中的数据文件</p></div><span className="v2-muted">{p.datasets.length} 个工作表</span></div>
      {p.datasets.length ? <div className="v2-dataset-list">{p.datasets.map(d=><button key={d.id} className="v2-dataset-row" onClick={()=>{p.setActiveId(d.id);location.hash='#/data'}}><span className="v2-file-icon"><FileSpreadsheet size={18}/></span><span className="v2-dataset-name"><b>{d.fileName}</b><small>{d.sheetName} · {d.headers.length} 个字段 · {d.rows.length.toLocaleString()} 行</small></span><MoreHorizontal size={17}/></button>)}</div> : <div className="v2-empty"><Database size={25}/><b>还没有导入数据</b><span>导入 Excel 或 CSV 后，这里会显示你的数据集。</span></div>}
    </section>
    <div className="v2-feature-grid">{[{icon:<Table2/>,title:'数据浏览',desc:'查看字段、搜索和筛选数据',to:'#/data'},{icon:<WandSparkles/>,title:'智能处理',desc:'清洗、转换并记录处理步骤',to:'#/transform'},{icon:<Combine/>,title:'多表合并',desc:'追加或关联不同数据来源',to:'#/merge'},{icon:<Bot/>,title:'AI 数据助手',desc:'自然语言生成可检查的操作计划',to:'#/transform'}].map(x=><a className="v2-feature" href={x.to} key={x.title}><span>{x.icon}</span><b>{x.title}</b><small>{x.desc}</small></a>)}</div>
  </>
}
function Stat({label,value,icon}:{label:string;value:string;icon:ReactNode}){return <div className="v2-stat"><span>{icon}</span><div><small>{label}</small><b>{value}</b></div></div>}
function QualityPanel({dataset,onExport}:{dataset:Dataset;onExport:()=>void}) {
  const quality = useMemo(() => inspectQuality(dataset), [dataset])
  return <div className="v2-quality-panel">
    <div className="v2-quality-head"><div><b>数据质量检查</b><small>自动识别空值、重复记录与字段类型</small></div><button className="v2-secondary" onClick={onExport}><Download size={14}/> 导出质量报告</button></div>
    <div className="v2-quality-stats">
      <div><small>数据完整度</small><b>{quality.completeness}%</b></div>
      <div><small>空值单元格</small><b>{quality.missingCells.toLocaleString()}</b></div>
      <div><small>完全重复行</small><b>{quality.duplicateRows.toLocaleString()}</b></div>
      <div><small>字段数量</small><b>{dataset.headers.length}</b></div>
    </div>
    <div className="v2-quality-table-wrap"><table className="v2-quality-table"><thead><tr><th>字段名称</th><th>识别类型</th><th>非空</th><th>空值</th><th>唯一值</th><th>完整度</th></tr></thead><tbody>{quality.fields.map(field=><tr key={field.name}><td>{field.name}</td><td><span className={'v2-type-tag '+(field.type==='混合类型'?'warning':'')}>{field.type}</span></td><td>{field.filled.toLocaleString()}</td><td>{field.missing.toLocaleString()}</td><td>{field.unique.toLocaleString()}</td><td><span className="v2-quality-meter"><i style={{width:field.completeness+'%'}}/></span>{field.completeness}%</td></tr>)}</tbody></table></div>
    {quality.fields.some(field=>field.type==='混合类型'||field.missing>0)&&<p className="v2-quality-hint"><ShieldCheck size={14}/> 检测到空值或混合类型字段，建议在处理前检查这些字段。</p>}
  </div>
}

function DataBrowser(p:Context){
  const [q,setQ]=useState('')
  const [sort,setSort]=useState<{i:number;dir:1|-1}|null>(null)
  const d=p.current
  const shown=useMemo(()=>{if(!d)return[];let rows=d.rows.map((r,i)=>({r,i})).filter(x=>!q||x.r.some(v=>cellText(v).toLowerCase().includes(q.toLowerCase())));if(sort)rows.sort((a,b)=>cellText(a.r[sort.i]).localeCompare(cellText(b.r[sort.i]),'zh-CN',{numeric:true})*sort.dir);return rows},[d,q,sort])
  return <><PageHeading eyebrow="DATA EXPLORER" title="数据浏览" description="搜索、排序并检查导入的数据结构。" action={<button className="v2-primary" onClick={p.openImport}><FilePlus2 size={15}/> 导入文件</button>}/>
    {!d?<EmptyPrompt onImport={p.openImport}/>:<section className="v2-card"><div className="v2-section-head"><div><h2>{d.fileName} / {d.sheetName}</h2><p>{d.rows.length.toLocaleString()} 行 · {d.headers.length} 个字段</p></div><button className="v2-secondary" onClick={p.exportCurrent}><Download size={15}/> 导出 Excel</button></div><QualityPanel dataset={d} onExport={() => exportQualityReport(d)}/><div className="v2-table-toolbar"><div className="v2-search"><Search size={15}/><input value={q} onChange={e=>setQ(e.target.value)} placeholder="搜索所有字段…"/></div><span>显示 {Math.min(500,shown.length).toLocaleString()} / {shown.length.toLocaleString()} 行</span></div><div className="v2-table-wrap"><table className="v2-table"><thead><tr><th>#</th>{d.headers.map((h,i)=><th key={i}><button onClick={()=>setSort(s=>s?.i===i?{i,dir:s.dir===1?-1:1}:{i,dir:1})}>{h}{sort?.i===i?(sort.dir===1?<ArrowUp size={12}/>:<ArrowDown size={12}/>):<ChevronDown size={12}/>}</button></th>)}</tr></thead><tbody>{shown.slice(0,500).map(({r,i})=><tr key={i}><td>{i+1}</td>{d.headers.map((_,j)=><td key={j}>{blank(r[j])?<em>空值</em>:cellText(r[j])}</td>)}</tr>)}</tbody></table></div></section>}
  </>
}
function EmptyPrompt({onImport}:{onImport:()=>void}){return <div className="v2-card v2-empty-prompt"><FileSpreadsheet size={32}/><h2>先导入一个数据文件</h2><p>支持 XLSX、XLS 和 CSV 格式。</p><button className="v2-primary" onClick={onImport}><FilePlus2 size={15}/> 选择文件</button></div>}
function Transform(p:Context){
  const [column,setColumn]=useState('')
  const d=p.current
  function applyTrim(){if(!d)return;const index=d.headers.indexOf(column);if(index<0)return;p.setDatasets(prev=>prev.map(x=>x.id===d.id?{...x,rows:x.rows.map(row=>row.map((v,i)=>i===index&&typeof v==='string'?v.trim():v))}:x));const step={id:uid(),label:'清理首尾空格',detail:'字段：'+column};p.setSteps(prev=>[...prev,step]);p.record('清理字段首尾空格：'+column)}
  function removeDuplicates(){if(!d)return;const next=[...new Map(d.rows.map(row=>[JSON.stringify(row),row])).values()];p.setDatasets(prev=>prev.map(x=>x.id===d.id?{...x,rows:next}:x));p.setSteps(prev=>[...prev,{id:uid(),label:'删除完全重复行',detail:'保留首次出现的记录'}]);p.record('删除完全重复行')}
  return <><PageHeading eyebrow="TRANSFORM STUDIO" title="智能数据处理" description="把每个处理动作变成可检查、可追踪的步骤。"/>
    {!d?<EmptyPrompt onImport={p.openImport}/>:<div className="v2-transform-layout"><section className="v2-card"><div className="v2-section-head"><div><h2>选择数据集</h2><p>{d.fileName} / {d.sheetName}</p></div></div><label className="v2-field-label">处理字段</label><select className="v2-select" value={column} onChange={e=>setColumn(e.target.value)}><option value="">选择一个字段</option>{d.headers.map(h=><option key={h}>{h}</option>)}</select><div className="v2-operation-list"><button disabled={!column} onClick={applyTrim}><WandSparkles size={17}/><span><b>清理首尾空格</b><small>删除所选文本字段前后的空格</small></span></button><button onClick={removeDuplicates}><Table2 size={17}/><span><b>删除完全重复行</b><small>按整行内容去重，保留首次出现项</small></span></button><button onClick={()=>{if(!d)return;const before=d.rows.length;const next=d.rows.filter(row=>row.some(value=>!blank(value)));const removed=before-next.length;p.setDatasets(prev=>prev.map(x=>x.id===d.id?{...x,rows:next}:x));p.setSteps(prev=>[...prev,{id:uid(),label:'删除全空行',detail:'删除 '+removed+' 行完全为空的记录'}]);p.record('删除全空行：'+removed+' 行')}}><Table2 size={17}/><span><b>删除全空行</b><small>移除所有字段均为空的记录</small></span></button></div><p className="v2-note"><ShieldCheck size={15}/> 当前操作修改浏览器中的工作副本，不会覆盖原始文件。导出前请检查结果。</p></section><section className="v2-card"><div className="v2-section-head"><div><h2>应用步骤</h2><p>按顺序记录本次会话的处理动作</p></div><span className="v2-muted">{p.steps.length} 步</span></div>{p.steps.length?<ol className="v2-steps">{p.steps.map((s,i)=><li key={s.id}><span>{i+1}</span><div><b>{s.label}</b><small>{s.detail}</small></div><Check size={15}/></li>)}</ol>:<div className="v2-empty"><Workflow size={25}/><b>还没有处理步骤</b><span>执行左侧操作后，步骤会出现在这里。</span></div>}<button className="v2-secondary v2-export-wide" onClick={p.exportCurrent}><Download size={15}/> 导出当前数据</button></section></div>}
  </>
}
function ModulePage({icon,title,description,bullets,status}:{icon:ReactNode;title:string;description:string;bullets:string[];status:string}){return <><PageHeading eyebrow="WORKSPACE MODULE" title={title} description={description}/><section className="v2-card v2-module-card"><div className="v2-module-icon">{icon}</div><div><span className="v2-status">{status}</span><h2>{title}能力规划</h2><p>该模块已纳入 V2.0 统一架构，当前先完成基础工作台和数据引擎，再接入完整交互。</p></div><ul>{bullets.map(x=><li key={x}>{x}</li>)}</ul></section></>}
function ActivityPage({history}:{history:string[]}){return <><PageHeading eyebrow="ACTIVITY LOG" title="操作记录" description="查看当前会话中执行过的数据操作。"/><section className="v2-card">{history.length?<div className="v2-steps">{history.map((x,i)=><div className="v2-activity-row" key={i}><span>{history.length-i}</span><b>{x}</b><Check size={15}/></div>)}</div>:<div className="v2-empty"><History size={25}/><b>暂无操作记录</b><span>数据处理操作会记录在这里。</span></div>}</section></>}
function SettingsPage({mode,setMode}:{mode:Mode;setMode:(m:Mode)=>void}){return <><PageHeading eyebrow="PREFERENCES & PRIVACY" title="设置与隐私" description="管理工作模式和数据共享偏好。"/><section className="v2-card v2-settings"><h2>工作模式</h2><p>简单模式提供引导式操作；专业模式展示更完整的配置与信息。</p><div className="v2-mode-switch large"><button className={mode==='simple'?'selected':''} onClick={()=>setMode('simple')}>简单模式</button><button className={mode==='professional'?'selected':''} onClick={()=>setMode('professional')}>专业模式</button></div><hr/><h2>数据隐私</h2><p>默认在本地浏览器处理文件。云端 AI 功能上线后，会在每次需要发送数据前说明发送内容、服务提供方和用途，并要求你明确授权。当前版本尚未启用云端数据传输。</p><div className="v2-privacy-banner"><ShieldCheck size={18}/> 本地优先 · 云端 AI 需明确授权</div></section></>}
function Root(){return <HashRouter><AppShell/></HashRouter>}
export default Root
