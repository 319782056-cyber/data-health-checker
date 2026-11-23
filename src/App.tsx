import { useMemo, useRef, useState, type ChangeEvent } from 'react'
import * as XLSX from 'xlsx'
import {
  Activity, ArrowDownUp, BarChart3, Check, ChevronDown, Download, FilePlus2,
  FileSpreadsheet, Filter, HardDrive, History, LayoutDashboard, ListFilter,
  Search, ShieldCheck, Sparkles, Table2, Trash2, WandSparkles, X, RotateCcw,
  Columns3, Rows3, Combine, CircleHelp
} from 'lucide-react'

type Cell = string | number | boolean | Date | null
type Dataset = { id: string; fileName: string; sheetName: string; headers: string[]; rows: Cell[][] }
type FilterRule = { column: string; value: string }
type HistoryItem = { label: string; before: Dataset[] }
const uid = () => Math.random().toString(36).slice(2, 10)
const cleanText = (v: Cell) => v instanceof Date ? v.toISOString().slice(0,10) : String(v ?? '')
const isBlank = (v: Cell) => v === null || v === undefined || String(v).trim() === ''

function App() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [datasets, setDatasets] = useState<Dataset[]>([])
  const [activeId, setActiveId] = useState('')
  const [query, setQuery] = useState('')
  const [filters, setFilters] = useState<FilterRule[]>([])
  const [sort, setSort] = useState<{col:number;dir:'asc'|'desc'}|null>(null)
  const [history, setHistory] = useState<HistoryItem[]>([])
  const [tab, setTab] = useState<'table'|'clean'|'pivot'>('table')
  const [pivotRow, setPivotRow] = useState('')
  const [pivotCol, setPivotCol] = useState('')
  const [pivotValue, setPivotValue] = useState('')
  const [pivotAgg, setPivotAgg] = useState<'sum'|'count'|'average'>('sum')
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const current = datasets.find(d => d.id === activeId)
  const headers = current?.headers ?? []
  const rows = current?.rows ?? []

  const filteredRows = useMemo(() => {
    let result = rows.map((row,index)=>({row,index})).filter(({row}) => {
      const q = query.trim().toLowerCase()
      if(q && !row.some(v=>cleanText(v).toLowerCase().includes(q))) return false
      return filters.every(f => {
        if(!f.value.trim()) return true
        const i=headers.indexOf(f.column)
        return i < 0 || cleanText(row[i]).toLowerCase().includes(f.value.toLowerCase())
      })
    })
    if(sort) result = [...result].sort((a,b)=>{
      const av=a.row[sort.col], bv=b.row[sort.col]
      const an=Number(av), bn=Number(bv)
      const bothNumeric=String(av??'').trim()!=='' && String(bv??'').trim()!=='' && Number.isFinite(an) && Number.isFinite(bn)
      const cmp=bothNumeric ? an-bn : cleanText(av).localeCompare(cleanText(bv),'zh-CN',{numeric:true,sensitivity:'base'})
      return sort.dir==='asc'?cmp:-cmp
    })
    return result
  },[rows,query,filters,headers,sort])

  const pivot = useMemo(() => {
    if(!current || !pivotRow) return []
    const ri=headers.indexOf(pivotRow), ci=pivotCol?headers.indexOf(pivotCol):-1, vi=pivotValue?headers.indexOf(pivotValue):-1
    const groups=new Map<string,Map<string,number[]>>()
    rows.forEach(row=>{
      const r=cleanText(row[ri])||'(空白)', c=ci<0?'总计':cleanText(row[ci])||'(空白)'
      if(!groups.has(r)) groups.set(r,new Map())
      const inner=groups.get(r)!
      if(!inner.has(c)) inner.set(c,[])
      const n=vi<0?1:Number(row[vi])
      if(pivotAgg==='count') inner.get(c)!.push(1)
      else if(Number.isFinite(n)) inner.get(c)!.push(n)
    })
    const cols=[...new Set([...groups.values()].flatMap(m=>[...m.keys()]))]
    return [...groups.entries()].map(([name,values])=>{
      const cells=cols.map(c=>{const arr=values.get(c)??[];return pivotAgg==='count'?arr.length:pivotAgg==='average'?(arr.length?arr.reduce((a,b)=>a+b,0)/arr.length:0):arr.reduce((a,b)=>a+b,0)})
      return {name,cells,total:cells.reduce((a,b)=>a+b,0)}
    }).concat([{name:'总计',cells:cols.map((_,i)=>[...groups.values()].reduce((s,m)=>{const a=m.get(cols[i])??[];return s+(pivotAgg==='count'?a.length:pivotAgg==='average'?(a.length?a.reduce((x,y)=>x+y,0)/a.length:0):a.reduce((x,y)=>x+y,0))},0)),total:0}])
  },[current,headers,rows,pivotRow,pivotCol,pivotValue,pivotAgg])

  function commit(label:string, transform:(d:Dataset)=>Dataset) {
    if(!current) return
    setHistory(h=>[...h.slice(-19),{label,before:datasets}])
    setDatasets(ds=>ds.map(d=>d.id===activeId?transform(d):d))
    setMessage(label+'，已完成；可撤销。')
  }
  async function importFiles(files:FileList|null) {
    if(!files?.length) return
    setBusy(true);setMessage('')
    try {
      const added:Dataset[]=[]
      for(const file of Array.from(files)){
        const ext=file.name.split('.').pop()?.toLowerCase()
        if(!['xlsx','xls','csv'].includes(ext??'')) continue
        const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true})
        wb.SheetNames.forEach(name=>{
          const raw=XLSX.utils.sheet_to_json<Cell[]>(wb.Sheets[name],{header:1,defval:null,raw:true})
          if(!raw.length) return
          const width=Math.max(...raw.map(r=>r.length))
          const hdr=Array.from({length:width},(_,i)=>cleanText(raw[0][i]).trim()||'字段 '+(i+1))
          added.push({id:uid(),fileName:file.name,sheetName:name,headers:hdr,rows:raw.slice(1).map(r=>Array.from({length:width},(_,i)=>r[i]??null))})
        })
      }
      if(!added.length) throw new Error('没有读取到有效工作表，请检查文件格式和内容。')
      setDatasets(ds=>[...ds,...added]);setActiveId(added[0].id);setTab('table');setHistory([]);setQuery('');setFilters([]);setMessage('成功导入 '+added.length+' 个工作表。')
    } catch(e){setMessage(e instanceof Error?e.message:'文件读取失败。')}
    finally {setBusy(false);if(inputRef.current)inputRef.current.value=''}
  }
  function exportCurrent(format:'xlsx'|'csv') {
    if(!current) return
    const data=[headers,...rows].map(row=>row.map(v=>v instanceof Date? v.toISOString():v))
    if(format==='csv'){
      const csv=data.map(row=>row.map(v=>{const s=cleanText(v);return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s}).join(',')).join('\r\n')
      const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8;'})
      downloadBlob(blob,current.sheetName+'.csv')
    }else{
      const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(data),current.sheetName.slice(0,31)||'数据')
      XLSX.writeFile(wb,current.sheetName+'.xlsx')
    }
  }
  function exportPivot(){
    if(!pivot.length) return
    const cols=[...new Set(pivot.flatMap(p=>p.cells.map((_,i)=>i)))]
    const data=[[pivotRow,...cols.map(i=>'分组 '+(i+1)),'汇总'],...pivot.map(p=>[p.name,...p.cells,p.total])]
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(data),'数据透视')
    XLSX.writeFile(wb,'数据透视结果.xlsx')
  }
  function undo(){
    const last=history[history.length-1];if(!last)return
    setDatasets(last.before);setHistory(h=>h.slice(0,-1));setMessage('已撤销：'+last.label)
  }
  function addFilter(){
    if(headers.length)setFilters(f=>[...f,{column:headers[0],value:''}])
  }
  function cleanAction(kind:'trim'|'dedupe'|'empty'|'fill'|'upper'|'lower'){
    if(!current)return
    if(kind==='fill'){
      const col=window.prompt('请输入要填充的字段名称：',headers[0]);if(!col||!headers.includes(col))return
      const value=window.prompt('请输入填充值：','0');if(value===null)return
      commit('填充空值：'+col,d=>({...d,rows:d.rows.map(r=>isBlank(r[d.headers.indexOf(col)])?r.map((v,i)=>i===d.headers.indexOf(col)?value:v):r)}));return
    }
    const labels={trim:'清理文本首尾空格',dedupe:'删除完全重复行',empty:'删除全空行',upper:'文本转大写',lower:'文本转小写'}
    commit(labels[kind],d=>{
      let next=d.rows.map(r=>r.map(v=>{
        if(typeof v!=='string')return v
        if(kind==='trim')return v.trim()
        if(kind==='upper')return v.toLocaleUpperCase()
        if(kind==='lower')return v.toLocaleLowerCase()
        return v
      }))
      if(kind==='dedupe')next=[...new Map(next.map(r=>[JSON.stringify(r),r])).values()]
      if(kind==='empty')next=next.filter(r=>r.some(v=>!isBlank(v)))
      return {...d,rows:next}
    })
  }
  const completeness=rows.length&&headers.length?Math.round(rows.reduce((s,r)=>s+r.filter(v=>!isBlank(v)).length,0)/(rows.length*headers.length)*100):100
  const duplicateCount=rows.length-new Set(rows.map(r=>JSON.stringify(r))).size

  return <div className="workbench">
    <aside className="wb-sidebar">
      <div className="wb-brand"><div className="wb-logo"><Activity size={21}/></div><div><b>Data<span>Workbench</span></b><small>LOCAL DATA STUDIO</small></div></div>
      <div className="wb-label">工作空间</div>
      <div className="wb-nav active"><LayoutDashboard size={17}/> 数据工作台</div>
      <div className="wb-nav"><Table2 size={17}/> 数据表处理</div>
      <div className="wb-nav"><WandSparkles size={17}/> 批量数据清洗</div>
      <div className="wb-nav"><BarChart3 size={17}/> 数据透视</div>
      <div className="wb-nav"><History size={17}/> 操作记录 <span>{history.length}</span></div>
      <div className="wb-sidebar-bottom"><div className="wb-privacy"><ShieldCheck size={17}/><div><b>本地隐私模式</b><small>文件只在当前浏览器处理</small></div><Check size={14}/></div><small>Data Workbench · V1.0</small></div>
    </aside>
    <main className="wb-main">
      <header className="wb-topbar"><div>工作空间 <span>/</span> <b>数据工作台</b></div><div className="wb-top-actions"><span className="local-tag"><i/> 本地处理</span><button className="wb-undo" disabled={!history.length} onClick={undo}><RotateCcw size={15}/> 撤销操作</button><button className="wb-import" onClick={()=>inputRef.current?.click()}><FilePlus2 size={16}/> 导入文件</button><input ref={inputRef} hidden multiple type="file" accept=".xlsx,.xls,.csv" onChange={e=>void importFiles(e.target.files)}/></div></header>
      <div className="wb-content">
        <section className="wb-hero"><div><div className="wb-eyebrow"><Sparkles size={13}/> DATA PREPARATION & ANALYSIS</div><h1>让繁琐的数据处理，<em>变成简单的工作流。</em></h1><p>导入、清洗、筛选、透视与导出，一站式完成数据准备工作。</p></div><div className="wb-hero-icon"><Activity size={34}/></div></section>
        {!datasets.length?<section className="wb-empty-upload"><div className="wb-upload-icon"><FileSpreadsheet size={30}/></div><h2>创建你的数据工作区</h2><p>支持一次导入多个 Excel / CSV 文件及多个工作表</p><button onClick={()=>inputRef.current?.click()} className="wb-primary"><FilePlus2 size={16}/> 选择数据文件</button><div className="wb-formats">XLSX · XLS · CSV <span/> 浏览器本地解析 · 不上传数据</div>{busy&&<p>正在读取文件…</p>}</section>:<>
          <div className="wb-dataset-strip"><div className="wb-dataset-tabs">{datasets.map(d=><button key={d.id} className={d.id===activeId?'selected':''} onClick={()=>{setActiveId(d.id);setFilters([]);setQuery('')}}><FileSpreadsheet size={14}/>{d.fileName} / {d.sheetName}<X size={12} onClick={e=>{e.stopPropagation();setDatasets(ds=>ds.filter(x=>x.id!==d.id));if(activeId===d.id)setActiveId(datasets.find(x=>x.id!==d.id)?.id??'')}}/></button>)}</div><button className="wb-add-file" onClick={()=>inputRef.current?.click()}><FilePlus2 size={14}/> 添加文件</button></div>
          <div className="wb-stats"><div><span>数据记录</span><b>{rows.length.toLocaleString()} <small>行</small></b></div><div><span>字段数量</span><b>{headers.length} <small>列</small></b></div><div><span>数据完整率</span><b>{completeness}%</b></div><div><span>完全重复行</span><b>{duplicateCount.toLocaleString()} <small>行</small></b></div></div>
          <section className="wb-panel">
            <div className="wb-panel-head"><div className="wb-tabs"><button className={tab==='table'?'on':''} onClick={()=>setTab('table')}><Table2 size={15}/> 数据表</button><button className={tab==='clean'?'on':''} onClick={()=>setTab('clean')}><WandSparkles size={15}/> 数据清洗</button><button className={tab==='pivot'?'on':''} onClick={()=>setTab('pivot')}><BarChart3 size={15}/> 数据透视</button></div><div className="wb-export"><button onClick={()=>exportCurrent('csv')}><Download size={14}/> CSV</button><button onClick={()=>exportCurrent('xlsx')}><Download size={14}/> Excel</button></div></div>
            {tab==='table'&&<><div className="wb-toolbar"><div className="wb-search"><Search size={15}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="搜索所有字段…"/></div><button onClick={addFilter}><Filter size={14}/> 添加筛选条件</button><span>{filteredRows.length.toLocaleString()} / {rows.length.toLocaleString()} 行</span></div>
              {filters.length>0&&<div className="wb-filter-list">{filters.map((f,i)=><div key={i}><select value={f.column} onChange={e=>setFilters(fs=>fs.map((x,j)=>i===j?{...x,column:e.target.value}:x))}>{headers.map(h=><option key={h}>{h}</option>)}</select><input value={f.value} onChange={e=>setFilters(fs=>fs.map((x,j)=>i===j?{...x,value:e.target.value}:x))} placeholder="包含文本…"/><button onClick={()=>setFilters(fs=>fs.filter((_,j)=>j!==i))}><X size={14}/></button></div>)}<button onClick={()=>setFilters([])}>清除筛选</button></div>}
              <div className="wb-table-scroll"><table className="wb-table"><thead><tr><th>#</th>{headers.map((h,i)=><th key={i} onClick={()=>setSort(s=>s?.col===i?{col:i,dir:s.dir==='asc'?'desc':'asc'}:{col:i,dir:'asc'})}>{h}<ArrowDownUp size={11}/></th>)}</tr></thead><tbody>{filteredRows.slice(0,500).map(({row,index})=><tr key={index}><td>{index+1}</td>{headers.map((_,i)=><td key={i}>{isBlank(row[i])?<span className="wb-null">空值</span>:cleanText(row[i])}</td>)}</tr>)}</tbody></table>{filteredRows.length>500&&<div className="wb-limit">当前预览前 500 行；筛选和导出作用于全部匹配记录。</div>}</div></>}
            {tab==='clean'&&<div className="wb-clean-area"><div className="wb-clean-intro"><h3><WandSparkles size={17}/> 批量数据清洗</h3><p>选择处理动作，系统将在当前工作副本上批量执行。每次操作都可以撤销。</p></div><div className="wb-clean-grid">
              <button onClick={()=>cleanAction('trim')}><span><Sparkles size={18}/></span><b>清理首尾空格</b><small>去除文本字段前后多余空格</small></button>
              <button onClick={()=>cleanAction('dedupe')}><span><Rows3 size={18}/></span><b>删除完全重复行</b><small>保留首次出现的记录</small></button>
              <button onClick={()=>cleanAction('empty')}><span><Trash2 size={18}/></span><b>删除全空行</b><small>移除没有任何内容的记录</small></button>
              <button onClick={()=>cleanAction('fill')}><span><Columns3 size={18}/></span><b>填充指定字段空值</b><small>指定字段与填充值，谨慎执行</small></button>
              <button onClick={()=>cleanAction('upper')}><span><WandSparkles size={18}/></span><b>文本转大写</b><small>批量转换文本字段</small></button>
              <button onClick={()=>cleanAction('lower')}><span><WandSparkles size={18}/></span><b>文本转小写</b><small>批量转换文本字段</small></button>
            </div><div className="wb-safe-note"><ShieldCheck size={16}/> 原始文件不会被修改。处理结果只存在当前浏览器工作区，导出后才会生成新文件。</div></div>}
            {tab==='pivot'&&<div className="wb-pivot-area"><div className="wb-pivot-config"><div><label>行字段</label><select value={pivotRow} onChange={e=>setPivotRow(e.target.value)}><option value="">选择分组字段</option>{headers.map(h=><option key={h}>{h}</option>)}</select></div><div><label>列字段（可选）</label><select value={pivotCol} onChange={e=>setPivotCol(e.target.value)}><option value="">不拆分列</option>{headers.filter(h=>h!==pivotRow).map(h=><option key={h}>{h}</option>)}</select></div><div><label>数值字段</label><select value={pivotValue} onChange={e=>setPivotValue(e.target.value)}><option value="">按记录计数</option>{headers.map(h=><option key={h}>{h}</option>)}</select></div><div><label>汇总方式</label><select value={pivotAgg} onChange={e=>setPivotAgg(e.target.value as typeof pivotAgg)}><option value="sum">求和</option><option value="count">计数</option><option value="average">平均值</option></select></div></div>
              {pivotRow?<><div className="wb-pivot-result-head"><b>透视结果</b><button onClick={exportPivot}><Download size={14}/> 导出透视结果</button></div><div className="wb-table-scroll"><table className="wb-table"><thead><tr><th>{pivotRow}</th>{pivot[0]?.cells.map((_,i)=><th key={i}>{pivotCol||'汇总 '+(i+1)}</th>)}<th>汇总</th></tr></thead><tbody>{pivot.map((p,i)=><tr key={i}><td>{p.name}</td>{p.cells.map((n,j)=><td key={j}>{Number(n.toFixed(2)).toLocaleString()}</td>)}<td>{Number(p.total.toFixed(2)).toLocaleString()}</td></tr>)}</tbody></table></div></>:<div className="wb-pivot-empty"><BarChart3 size={27}/><b>配置字段，生成数据透视表</b><small>选择行字段后即可按类别汇总数据</small></div>}
            </div>}
          </section>
          <div className="wb-bottom-info"><span><ShieldCheck size={14}/> 文件内容仅在本地浏览器处理</span><span>当前工作区：{datasets.length} 个工作表</span></div>
        </>}
        {message&&<div className="wb-toast" role="status">{message}<button onClick={()=>setMessage('')}><X size={14}/></button></div>}
      </div>
    </main>
  </div>
}
function downloadBlob(blob:Blob,name:string){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();URL.revokeObjectURL(url)}
export default App
