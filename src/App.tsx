import { useMemo, useRef, useState, type ReactNode } from 'react'
import * as XLSX from 'xlsx'
import {
  Activity, AlertCircle, ArrowUpRight, Check, ChevronDown, CircleHelp,
  ClipboardList, CloudUpload, FileSpreadsheet, FileText, HardDrive,
  LayoutDashboard, RefreshCw, ShieldCheck, Sparkles, Table2, Trash2,
  
} from 'lucide-react'

type Cell = string | number | boolean | Date | null
type SheetData = { name: string; rows: Cell[][] }

function App() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [fileName, setFileName] = useState('')
  const [sheets, setSheets] = useState<SheetData[]>([])
  const [activeSheet, setActiveSheet] = useState('')
  const [dragging, setDragging] = useState(false)
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)

  const current = sheets.find((sheet) => sheet.name === activeSheet)
  const rows = current?.rows ?? []
  const headers = rows[0]?.map((value, index) => String(value ?? '').trim() || `字段 ${index + 1}`) ?? []
  const records = rows.slice(1)


  const columnProfiles = useMemo(() => {
    return headers.map((header, columnIndex) => {
      const values = records
        .map((row) => row[columnIndex])
        .filter((value) => value !== null && value !== undefined && String(value).trim() !== '')
      const normalized = values.map((value) => String(value).trim())
      const uniqueValues = new Set(normalized)
      const types = new Set(normalized.map(inferCellType))
      const missing = Math.max(0, records.length - values.length)
      const typeLabel = types.size === 0
        ? '暂无数据'
        : types.size > 1
          ? '混合类型'
          : ({ number: '数值', date: '日期', boolean: '布尔值', text: '文本' } as Record<string, string>)[[...types][0]] ?? '文本'
      return {
        name: header,
        missing,
        missingRate: records.length ? Math.round(missing / records.length * 100) : 0,
        unique: uniqueValues.size,
        repeated: Math.max(0, values.length - uniqueValues.size),
        typeLabel,
        samples: [...uniqueValues].slice(0, 3),
      }
    })
  }, [headers, records])

  const stats = useMemo(() => {
    if (!rows.length) return { rowCount: 0, colCount: 0, missing: 0, duplicate: 0, completeness: 100 }
    let missing = 0
    records.forEach((row) => {
      headers.forEach((_, index) => {
        if (row[index] === null || row[index] === undefined || String(row[index]).trim() === '') missing += 1
      })
    })
    const signatures = records.map((row) => JSON.stringify(row))
    const duplicate = signatures.length - new Set(signatures).size
    const total = records.length * headers.length
    return {
      rowCount: records.length,
      colCount: headers.length,
      missing,
      duplicate,
      completeness: total ? Math.round((total - missing) / total * 100) : 100,
    }
  }, [rows, records, headers])

  async function loadFile(file?: File) {
    if (!file) return
    setError('')
    const ext = file.name.split('.').pop()?.toLowerCase()
    if (!['xlsx', 'xls', 'csv'].includes(ext ?? '')) {
      setError('暂不支持该文件格式，请上传 XLSX、XLS 或 CSV 文件。')
      return
    }
    setBusy(true)
    try {
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
      const parsed = workbook.SheetNames.map((name) => ({
        name,
        rows: XLSX.utils.sheet_to_json<Cell[]>(workbook.Sheets[name], {
          header: 1,
          defval: null,
          raw: false,
        }),
      }))
      setSheets(parsed)
      setActiveSheet(parsed[0]?.name ?? '')
      setFileName(file.name)
    } catch {
      setError('文件读取失败，请确认文件未损坏或受密码保护。')
    } finally {
      setBusy(false)
    }
  }

  function reset() {
    setFileName('')
    setSheets([])
    setActiveSheet('')
    setError('')
    if (inputRef.current) inputRef.current.value = ''
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark"><Activity size={21} strokeWidth={2.6} /></div>
          <div><strong>Data<span>Care</span></strong><small>DATA HEALTH CHECKER</small></div>
        </div>
        <div className="workspace-label">工作空间</div>
        <nav>
          <a className="nav-item active"><LayoutDashboard size={18} /> 数据概览</a>
          <a className="nav-item"><ShieldCheck size={18} /> 数据体检 <span className="soon">当前</span></a>
          <a className="nav-item muted"><Sparkles size={18} /> 数据清洗 <span className="lock">即将推出</span></a>
          <a className="nav-item muted"><ClipboardList size={18} /> 处理历史</a>
        </nav>
        <div className="sidebar-bottom">
          <div className="privacy-card"><div className="privacy-icon"><HardDrive size={17}/></div><div><b>本地隐私保护</b><p>文件只在你的浏览器内读取和分析，不会上传到应用服务器。</p></div><Check size={15} className="privacy-check"/></div>
          <div className="profile"><div className="avatar">DH</div><div><b>个人工作区</b><small>免费使用 · 无需登录</small></div><CircleHelp size={17} className="help-icon"/></div>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div className="breadcrumb">工作空间 <span>/</span> <b>数据体检</b></div>
          <div className="top-actions"><span className="local-pill"><span/> 本地处理模式</span><button className="icon-button" aria-label="帮助"><CircleHelp size={18}/></button><div className="avatar small-avatar">DH</div></div>
        </header>

        <div className="content">
          <section className="page-heading">
            <div><div className="eyebrow"><span className="eyebrow-dot"/> DATA QUALITY WORKSPACE</div><h1>让数据更健康，<em>从一次体检开始。</em></h1><p>快速发现 Excel 和 CSV 中的潜在问题，让每一份数据都清晰、可靠、可用。</p></div>
            <div className="heading-art"><div className="art-orbit orbit-one"/><div className="art-orbit orbit-two"/><div className="art-core"><Activity size={34}/></div><span className="art-spark spark-a">✦</span><span className="art-spark spark-b">✧</span></div>
          </section>

          {!fileName ? (
            <>
              <section className="upload-panel">
                <div className="section-top"><div><h2>开始数据体检</h2><p>上传文件，系统将自动读取并分析数据质量。</p></div><span className="step-tag"><span>01</span> 上传数据</span></div>
                <div className={`dropzone ${dragging ? 'dragging' : ''}`}
                  onDragOver={(event) => { event.preventDefault(); setDragging(true) }}
                  onDragLeave={() => setDragging(false)}
                  onDrop={(event) => { event.preventDefault(); setDragging(false); void loadFile(event.dataTransfer.files[0]) }}>
                  <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(event) => void loadFile(event.target.files?.[0])}/>
                  <div className="upload-icon"><CloudUpload size={29}/></div>
                  <h3>拖拽文件到这里，或 <button className="text-button" onClick={() => inputRef.current?.click()}>点击上传</button></h3>
                  <p>支持 Excel 工作簿和 CSV 文件，单个文件最大建议 20 MB</p>
                  <div className="privacy-note"><ShieldCheck size={15}/><span><b>隐私说明：</b>文件内容仅在当前浏览器处理，不会上传到应用服务器或发送给第三方。点击“移除文件”可清除当前工作区中的数据。</span></div>
                  <div className="file-types"><span><FileSpreadsheet size={15}/> XLSX / XLS</span><span><FileText size={15}/> CSV</span><span><ShieldCheck size={15}/> 安全本地解析</span></div>
                </div>
                {error && <div className="error-message"><AlertCircle size={16}/>{error}</div>}
                {busy && <div className="loading-line"><RefreshCw size={15} className="spin"/> 正在读取文件…</div>}
              </section>

              <section className="features-section"><div className="section-title-row"><div><h2>一次体检，全面了解数据质量</h2><p>从完整性、唯一性、有效性和一致性四个维度发现问题。</p></div><span className="light-label">核心检测能力</span></div>
                <div className="feature-grid">
                  <div className="feature-card"><div className="feature-icon blue"><Table2 size={19}/></div><b>缺失值检测</b><p>识别空单元格、空字符串和字段缺失。</p><span>COMPLETENESS <ArrowUpRight size={13}/></span></div>
                  <div className="feature-card"><div className="feature-icon violet"><CopyIcon/></div><b>重复数据检测</b><p>发现完全重复记录，辅助识别重复主键。</p><span>UNIQUENESS <ArrowUpRight size={13}/></span></div>
                  <div className="feature-card"><div className="feature-icon amber"><AlertCircle size={19}/></div><b>类型与异常检测</b><p>识别类型混杂、格式异常和可疑数值。</p><span>VALIDITY <ArrowUpRight size={13}/></span></div>
                  <div className="feature-card"><div className="feature-icon green"><RefreshCw size={19}/></div><b>字段一致性</b><p>检查空格、大小写和文本格式差异。</p><span>CONSISTENCY <ArrowUpRight size={13}/></span></div>
                </div>
              </section>
              <footer className="footer"><span>© 2026 DataCare · Data Health Checker</span><span><ShieldCheck size={14}/> 数据安全优先 <i/> 当前版本 v0.1.0</span></footer>
            </>
          ) : (
            <>
              <section className="file-banner"><div className="file-symbol"><FileSpreadsheet size={22}/></div><div className="file-meta"><b>{fileName}</b><small>{sheets.length} 个工作表 · 已在浏览器本地读取</small></div><button className="outline-button" onClick={reset}><Trash2 size={15}/> 移除文件</button></section>
              <div className="metrics-grid">
                <Metric label="数据记录" value={stats.rowCount.toLocaleString()} suffix="行" icon={<Table2/>} tone="blue"/>
                <Metric label="字段数量" value={String(stats.colCount)} suffix="列" icon={<FileSpreadsheet/>} tone="violet"/>
                <Metric label="缺失单元格" value={stats.missing.toLocaleString()} suffix="个" icon={<AlertCircle/>} tone="amber"/>
                <Metric label="完全重复行" value={stats.duplicate.toLocaleString()} suffix="行" icon={<RefreshCw/>} tone="green"/>
              </div>
              <section className="profile-panel">
                <div className="profile-heading">
                  <div><h2>字段级质量分析</h2><p>逐列查看缺失率、唯一值、重复值和推断类型</p></div>
                  <span className="profile-count">{columnProfiles.length} 个字段</span>
                </div>
                {columnProfiles.length ? (
                  <div className="profile-table-wrap">
                    <table className="profile-table">
                      <thead><tr><th>字段名称</th><th>推断类型</th><th>缺失值</th><th>缺失率</th><th>唯一值</th><th>重复值</th><th>示例值</th></tr></thead>
                      <tbody>{columnProfiles.map((column, index) => (
                        <tr key={index}>
                          <td><b>{column.name}</b></td>
                          <td><span className={`type-chip ${column.typeLabel === '混合类型' ? 'mixed' : ''}`}>{column.typeLabel}</span></td>
                          <td>{column.missing.toLocaleString()}</td>
                          <td><div className="missing-cell"><span>{column.missingRate}%</span><div className="missing-track"><i style={{ width: `${column.missingRate}%` }} /></div></div></td>
                          <td>{column.unique.toLocaleString()}</td>
                          <td>{column.repeated.toLocaleString()}</td>
                          <td><div className="sample-values">{column.samples.length ? column.samples.map((sample, sampleIndex) => <span key={sampleIndex} title={sample}>{sample}</span>) : <span className="no-sample">暂无</span>}</div></td>
                        </tr>
                      ))}</tbody>
                    </table>
                  </div>
                ) : <div className="empty-state">没有可分析的字段。</div>}
              </section>
              <section className="data-panel"><div className="data-panel-head"><div><h2>数据预览</h2><p>已读取数据 · 当前展示前 100 条记录</p></div><div className="sheet-select-wrap"><label>工作表</label><div className="select-shell"><select value={activeSheet} onChange={(event) => setActiveSheet(event.target.value)}>{sheets.map((sheet) => <option key={sheet.name}>{sheet.name}</option>)}</select><ChevronDown size={15}/></div></div></div>
                <div className="table-wrap"><table><thead><tr><th className="row-num">#</th>{headers.map((header, index) => <th key={index}>{header}</th>)}</tr></thead><tbody>{records.slice(0,100).map((row, rowIndex) => <tr key={rowIndex}><td className="row-num">{rowIndex + 1}</td>{headers.map((_, colIndex) => <td key={colIndex}>{row[colIndex] === null || row[colIndex] === undefined || row[colIndex] === '' ? <span className="empty-cell">空值</span> : String(row[colIndex])}</td>)}</tr>)}</tbody></table>{!records.length && <div className="empty-state">当前工作表没有可预览的数据行。</div>}</div>
                <div className="table-foot"><span>共 {stats.rowCount.toLocaleString()} 条记录</span><span>完整性 <b>{stats.completeness}%</b></span></div>
              </section>
              <div className="next-step"><div><b><Sparkles size={17}/> 数据读取完成</b><p>基础体检指标已生成。更深入的字段分析与清洗操作将在后续版本开放。</p></div><button className="primary-button" disabled>深入体检 <ArrowUpRight size={16}/></button></div>
              <footer className="footer"><span>© 2026 DataCare · Data Health Checker</span><span><ShieldCheck size={14}/> 数据仅在本地处理 <i/> v0.1.0</span></footer>
            </>
          )}
        </div>
      </main>
    </div>
  )
}

function inferCellType(value: Cell): 'number' | 'date' | 'boolean' | 'text' {
  if (typeof value === 'number') return 'number'
  if (typeof value === 'boolean') return 'boolean'
  if (value instanceof Date) return 'date'
  const text = String(value ?? '').trim()
  if (/^(true|false|是|否)$/i.test(text)) return 'boolean'
  if (/^[+-]?(?:\\d+\\.?\\d*|\\.\\d+)(?:e[+-]?\\d+)?%?$/i.test(text.replace(/,/g, ''))) return 'number'
  if (/^\\d{4}[-/.年]\\d{1,2}[-/.月]\\d{1,2}日?$/.test(text)) return 'date'
  return 'text'
}

function CopyIcon() { return <ClipboardList size={19}/> }
function Metric({ label, value, suffix, icon, tone }: { label: string; value: string; suffix: string; icon: ReactNode; tone: string }) {
  return <div className="metric-card"><div className={`metric-icon ${tone}`}>{icon}</div><div><span>{label}</span><strong>{value}<small>{suffix}</small></strong></div></div>
}

export default App
