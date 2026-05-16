import { useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent, FormEvent, ReactNode } from 'react'
import {
  Activity,
  BarChart3,
  ClipboardCheck,
  Drill,
  Download,
  Edit3,
  FileDown,
  HardHat,
  History,
  LayoutDashboard,
  Menu,
  MoreVertical,
  Mountain,
  Pickaxe,
  RefreshCw,
  Save,
  ShieldCheck,
  Trash2,
  Truck,
  Upload,
  X,
} from 'lucide-react'
import { Capacitor } from '@capacitor/core'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import {
  jumboEquipoOptions,
  retroEquipoOptions,
  scoopEquipoOptions,
} from './data/equipos'
import './App.css'

type Shift = '1' | '2'
type RecordType = 'barrenacion' | 'rezagado' | 'seguridad'
type BarrenacionActivity = 'jumbo' | 'maquinaPierna' | 'voladura'
type RezagadoEquipment = 'scoop' | 'retro'
type HaulActivity = 'rezagado' | 'traspaleo' | 'limpia' | 'balastreo' | 'planilla' | 'relleno'
type RetroActivity = 'amacice' | 'reAmacice' | 'tableo' | 'limpia' | 'balastreo' | 'mtAcequia'

type DrillRow = {
  equipo: string
  operador: string
  ayudante: string
  nivelObra: string
  rpaCfte: string
  longitud: number
  cuele: number
  desarrollo: number
  barrenosDados: number
  barrenosCargados: number
  metrosDados: number
  horasServicio: number
  zanco: number
  cople: number
  barra: number
  broca: number
  horometroDieselInicial: number
  horometroDieselFinal: number
  horometroElectInicial: number
  horometroElectFinal: number
}

type BlastRow = {
  obra: string
  oficial: string
  ayudante: string
  rpaCfte: string
  longitud: number
  cuele: number
  desarrollo: number
  barrenosPegados: number
  metrosPegados: number
  horasServicio: number
  anfoInicial: number
  anfoFinal: number
  observaciones: string
}

type HaulRow = {
  equipo: string
  operador: string
  nivelObra: string
  destino: string
  rezagado: number
  traspaleo: number
  limpia: number
  balastreo: number
  planilla: number
  relleno: number
  camiones: number
  horometroInicial: number
  horometroFinal: number
  diesel: number
  observaciones: string
}

type RetroRow = {
  equipo: string
  operador: string
  nivelObra: string
  amacice: number
  reAmacice: number
  tableo: number
  limpia: number
  balastreo: number
  mtAcequia: number
  horometroInicial: number
  horometroFinal: number
  diesel: number
  observaciones: string
}

type BaseRecord = {
  id: string
  type: RecordType
  supervisor: string
  fecha: string
  turno: Shift
  unidad: string
  createdAt: string
  updatedAt: string
  deletedAt?: string
  syncedAt?: string
}

type BarrenacionRecord = BaseRecord & {
  type: 'barrenacion'
  jumbo: DrillRow[]
  maquinaPierna: DrillRow[]
  voladuras: BlastRow[]
  polvorero: string
  choferCamion: string
  choferPipa: string
  bobCat: string
  bombeo: string
  servicios: string
  comentarios: string
  inasistencias: string
}

type RezagadoRecord = BaseRecord & {
  type: 'rezagado'
  scoopTram: HaulRow[]
  retro: RetroRow[]
  comentarios: string
}

type SeguridadRecord = BaseRecord & {
  type: 'seguridad'
  accidentes: number
  incidentes: number
  actosInseguros: string
  condicionesInseguras: string
  fuerzaLaboral: number
  platicaSeguridad: string
  actividadesSeguridad: string
  correccionesMejoras: string
  actividadesOperacion: string
  observaciones: string
}

type MineRecord = BarrenacionRecord | RezagadoRecord | SeguridadRecord

const STORAGE_KEY = 'mga-bitacora-operaciones-v1'
const API_URL_KEY = 'mga-bitacora-api-url'
const DEFAULT_API_URL = 'https://mga-bitacora-mina.onrender.com'
const today = new Date().toISOString().slice(0, 10)

const baseDefaults = {
  supervisor: '',
  fecha: today,
  turno: '1' as Shift,
  unidad: 'Unidad Providencia',
}

const nowIso = () => new Date().toISOString()
const defaultJumbo = jumboEquipoOptions[0] ?? 'JUMBO'
const defaultScoop = scoopEquipoOptions[0] ?? 'SCOOP TRAM'
const defaultRetro = retroEquipoOptions[0] ?? 'RETRO'

const emptyDrillRow = (equipo: string): DrillRow => ({
  equipo,
  operador: '',
  ayudante: '',
  nivelObra: '',
  rpaCfte: '',
  longitud: 0,
  cuele: 0,
  desarrollo: 0,
  barrenosDados: 0,
  barrenosCargados: 0,
  metrosDados: 0,
  horasServicio: 0,
  zanco: 0,
  cople: 0,
  barra: 0,
  broca: 0,
  horometroDieselInicial: 0,
  horometroDieselFinal: 0,
  horometroElectInicial: 0,
  horometroElectFinal: 0,
})

const emptyBlastRow = (): BlastRow => ({
  obra: '',
  oficial: '',
  ayudante: '',
  rpaCfte: '',
  longitud: 0,
  cuele: 0,
  desarrollo: 0,
  barrenosPegados: 0,
  metrosPegados: 0,
  horasServicio: 0,
  anfoInicial: 0,
  anfoFinal: 0,
  observaciones: '',
})

const emptyHaulRow = (equipo: string): HaulRow => ({
  equipo,
  operador: '',
  nivelObra: '',
  destino: '',
  rezagado: 0,
  traspaleo: 0,
  limpia: 0,
  balastreo: 0,
  planilla: 0,
  relleno: 0,
  camiones: 0,
  horometroInicial: 0,
  horometroFinal: 0,
  diesel: 0,
  observaciones: '',
})

const emptyRetroRow = (): RetroRow => ({
  equipo: defaultRetro,
  operador: '',
  nivelObra: '',
  amacice: 0,
  reAmacice: 0,
  tableo: 0,
  limpia: 0,
  balastreo: 0,
  mtAcequia: 0,
  horometroInicial: 0,
  horometroFinal: 0,
  diesel: 0,
  observaciones: '',
})

const makeBarrenacion = (): BarrenacionRecord => {
  const now = nowIso()
  return {
    id: crypto.randomUUID(),
    type: 'barrenacion',
    createdAt: now,
    updatedAt: now,
    ...baseDefaults,
    jumbo: [emptyDrillRow(defaultJumbo)],
    maquinaPierna: [],
    voladuras: [],
    polvorero: '',
    choferCamion: '',
    choferPipa: '',
    bobCat: '',
    bombeo: '',
    servicios: '',
    comentarios: '',
    inasistencias: '',
  }
}

const makeRezagado = (): RezagadoRecord => {
  const now = nowIso()
  return {
    id: crypto.randomUUID(),
    type: 'rezagado',
    createdAt: now,
    updatedAt: now,
    ...baseDefaults,
    scoopTram: [emptyHaulRow(defaultScoop)],
    retro: [],
    comentarios: '',
  }
}

const makeSeguridad = (): SeguridadRecord => {
  const now = nowIso()
  return {
    id: crypto.randomUUID(),
    type: 'seguridad',
    createdAt: now,
    updatedAt: now,
    ...baseDefaults,
    accidentes: 0,
    incidentes: 0,
    actosInseguros: '',
    condicionesInseguras: '',
    fuerzaLaboral: 0,
    platicaSeguridad: '',
    actividadesSeguridad: '',
    correccionesMejoras: '',
    actividadesOperacion: '',
    observaciones: '',
  }
}

function App() {
  const [section, setSection] = useState<'captura' | 'dashboard'>(getInitialSection)
  const [formType, setFormType] = useState<RecordType>('barrenacion')
  const [barrenacion, setBarrenacion] = useState<BarrenacionRecord>(makeBarrenacion)
  const [rezagado, setRezagado] = useState<RezagadoRecord>(makeRezagado)
  const [seguridad, setSeguridad] = useState<SeguridadRecord>(makeSeguridad)
  const [records, setRecords] = useState<MineRecord[]>(loadRecords)
  const [query, setQuery] = useState('')
  const [message, setMessage] = useState('')
  const [editingId, setEditingId] = useState<string | null>(null)
  const [syncing, setSyncing] = useState(false)
  const [apiUrl, setApiUrl] = useState(loadApiUrl)
  const reportRef = useRef<HTMLDivElement>(null)

  const visibleRecords = useMemo(() => records.filter((record) => !record.deletedAt), [records])
  const filtered = useMemo(() => {
    const term = query.trim().toLowerCase()
    if (!term) return visibleRecords
    return visibleRecords.filter((record) =>
      [record.type, record.supervisor, record.fecha, record.turno, record.unidad]
        .join(' ')
        .toLowerCase()
        .includes(term),
    )
  }, [visibleRecords, query])
  const kpis = useMemo(() => computeKpis(filtered), [filtered])
  const recordCounts = useMemo(
    () => ({
      barrenacion: visibleRecords.filter((record) => record.type === 'barrenacion').length,
      rezagado: visibleRecords.filter((record) => record.type === 'rezagado').length,
      seguridad: visibleRecords.filter((record) => record.type === 'seguridad').length,
      total: visibleRecords.length,
    }),
    [visibleRecords],
  )
  const recentRecords = useMemo(() => visibleRecords.slice(0, 4), [visibleRecords])
  const pendingSync = useMemo(
    () => records.filter((record) => record.updatedAt !== record.syncedAt).length,
    [records],
  )

  useEffect(() => {
    localStorage.setItem(API_URL_KEY, apiUrl)
  }, [apiUrl])

  useEffect(() => {
    void syncRecords({ silent: true })
    const handleOnline = () => void syncRecords({ silent: true })
    window.addEventListener('online', handleOnline)
    return () => window.removeEventListener('online', handleOnline)
    // The initial sync must run once on startup; saves trigger their own sync with fresh records.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  function persist(next: MineRecord[]) {
    const normalized = mergeRecords(next.map(normalizeRecord))
    setRecords(normalized)
    localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized))
    return normalized
  }

  function saveRecord(event: FormEvent) {
    event.preventDefault()
    const source =
      formType === 'barrenacion' ? barrenacion : formType === 'rezagado' ? rezagado : seguridad
    const updated = normalizeRecord({
      ...source,
      updatedAt: nowIso(),
      deletedAt: undefined,
      syncedAt: undefined,
    } as MineRecord)
    const next = editingId
      ? records.map((record) => (record.id === editingId ? updated : record))
      : [updated, ...records]
    const persisted = persist(next)
    resetForm(formType)
    setEditingId(null)
    setMessage(editingId ? 'Captura actualizada offline.' : 'Registro guardado offline en este dispositivo.')
    void syncRecords({ silent: true, sourceRecords: persisted })
  }

  function resetForm(type: RecordType) {
    if (type === 'barrenacion') setBarrenacion(makeBarrenacion())
    if (type === 'rezagado') setRezagado(makeRezagado())
    if (type === 'seguridad') setSeguridad(makeSeguridad())
  }

  function startEdit(record: MineRecord) {
    const cleanRecord = { ...record, deletedAt: undefined } as MineRecord
    setFormType(record.type)
    if (record.type === 'barrenacion') setBarrenacion(cleanRecord as BarrenacionRecord)
    if (record.type === 'rezagado') setRezagado(cleanRecord as RezagadoRecord)
    if (record.type === 'seguridad') setSeguridad(cleanRecord as SeguridadRecord)
    setEditingId(record.id)
    setSection('captura')
    setMessage('Modo edicion activo. Guarda para conservar los cambios.')
  }

  function cancelEdit() {
    resetForm(formType)
    setEditingId(null)
    setMessage('Edicion cancelada.')
  }

  function selectFormType(type: RecordType) {
    if (editingId && type !== formType) {
      setMessage('Guarda o cancela la edicion antes de cambiar de formato.')
      return
    }
    setFormType(type)
  }

  function removeRecord(id: string) {
    const deletedAt = nowIso()
    const persisted = persist(
      records.map((record) =>
        record.id === id
          ? normalizeRecord({ ...record, deletedAt, updatedAt: deletedAt, syncedAt: undefined } as MineRecord)
          : record,
      ),
    )
    setMessage('Captura eliminada localmente.')
    void syncRecords({ silent: true, sourceRecords: persisted })
  }

  function exportJson() {
    const blob = new Blob([JSON.stringify(records, null, 2)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `bitacora-mga-${today}.json`
    link.click()
    URL.revokeObjectURL(url)
  }

  async function importJson(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return
    const imported = JSON.parse(await file.text()) as MineRecord[]
    const persisted = persist([...imported.map(normalizeRecord), ...records])
    setMessage('Datos importados al panel local.')
    void syncRecords({ silent: true, sourceRecords: persisted })
    event.target.value = ''
  }

  async function syncRecords(options: { silent?: boolean; sourceRecords?: MineRecord[] } = {}) {
    const apiBase = resolveApiBase(apiUrl)
    if (!apiBase || (typeof navigator !== 'undefined' && !navigator.onLine)) {
      if (!options.silent) setMessage('Sin internet. La captura queda guardada y se enviara automaticamente.')
      return
    }
    setSyncing(true)
    try {
      const localRecords = (options.sourceRecords ?? records).map(normalizeRecord)
      const response = await fetch(`${apiBase}/api/records/sync`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ records: localRecords }),
      })
      if (!response.ok) throw new Error(`HTTP ${response.status}`)
      const payload = (await response.json()) as { records: MineRecord[] }
      const syncedAt = nowIso()
      persist(
        mergeRecords([...localRecords, ...payload.records.map(normalizeRecord)]).map((record) => ({
          ...record,
          syncedAt: record.updatedAt ?? syncedAt,
        })),
      )
      if (!options.silent) setMessage('Sincronizacion completada con Render.')
    } catch {
      if (!options.silent) setMessage('No se pudo sincronizar. Se reintentara automaticamente cuando haya red.')
    } finally {
      setSyncing(false)
    }
  }

  async function exportPdf() {
    if (!reportRef.current) return
    const canvas = await html2canvas(reportRef.current, { scale: 2, backgroundColor: '#ffffff' })
    const img = canvas.toDataURL('image/png')
    const pdf = new jsPDF('p', 'mm', 'a4')
    const pageWidth = pdf.internal.pageSize.getWidth()
    const pageHeight = pdf.internal.pageSize.getHeight()
    const imgHeight = (canvas.height * pageWidth) / canvas.width
    let heightLeft = imgHeight
    let position = 0
    pdf.addImage(img, 'PNG', 0, position, pageWidth, imgHeight)
    heightLeft -= pageHeight
    while (heightLeft > 0) {
      position = heightLeft - imgHeight
      pdf.addPage()
      pdf.addImage(img, 'PNG', 0, position, pageWidth, imgHeight)
      heightLeft -= pageHeight
    }
    pdf.save(`reporte-kpi-mga-${today}.pdf`)
  }

  return (
    <main className={`app-shell ${section === 'captura' ? 'capture-mode' : 'review-mode'}`}>
      <header className="topbar">
        <button className="chrome-icon" aria-label="Menu" type="button">
          <Menu size={21} />
        </button>
        <div className="brand">
          <img src="/mga-logo.jfif" alt="MGA" />
          <div>
            <span>MGA Contratista Minera</span>
            <strong>Bitacora de Operaciones Mina</strong>
          </div>
        </div>
        <nav className="mode-tabs" aria-label="Vista principal">
          <button className={section === 'captura' ? 'active' : ''} onClick={() => setSection('captura')}>
            <ClipboardCheck size={18} /> Captura
          </button>
          <button
            className={section === 'dashboard' ? 'active' : ''}
            onClick={() => {
              setSection('dashboard')
              void syncRecords({ silent: true })
            }}
          >
            <LayoutDashboard size={18} /> Revision web
          </button>
        </nav>
        <button className="sync-button" onClick={() => void syncRecords()} disabled={syncing}>
          <RefreshCw size={18} className={syncing ? 'spin' : ''} />
          {syncing ? 'Conectando' : pendingSync > 0 ? 'Pendiente' : 'Al dia'}
          {pendingSync > 0 && <span>{pendingSync}</span>}
        </button>
        <button className="chrome-icon" aria-label="Mas opciones" type="button">
          <MoreVertical size={21} />
        </button>
      </header>

      {message && <div className="toast">{message}</div>}

      {section === 'captura' ? (
        <form className="workspace" onSubmit={saveRecord}>
          <aside className="side-panel">
            <div className="capture-badge">
              <Mountain size={18} /> Operacion mina
            </div>
            <h1>{editingId ? 'Editar captura' : 'Nueva captura'}</h1>
            <p>Captura un evento de campo con los datos clave. El reporte completo se arma en la revision web.</p>
            <div className="connection-card">
              <span>Conexion automatica</span>
              <strong>{syncing ? 'Sincronizando' : pendingSync > 0 ? `${pendingSync} por enviar` : 'Datos al dia'}</strong>
            </div>
            <div className="record-tabs">
              <button
                type="button"
                className={`module-card module-red ${formType === 'barrenacion' ? 'active' : ''}`}
                onClick={() => selectFormType('barrenacion')}
              >
                <span className="module-icon"><Drill size={24} /></span>
                <span className="module-title">Barrenos</span>
                <span className="module-subtitle">Voladuras</span>
                <span className="module-ring"><strong>{recordCounts.barrenacion}</strong><small>reg.</small></span>
              </button>
              <button
                type="button"
                className={`module-card module-blue ${formType === 'rezagado' ? 'active' : ''}`}
                onClick={() => selectFormType('rezagado')}
              >
                <span className="module-icon"><Truck size={24} /></span>
                <span className="module-title">Rezagado</span>
                <span className="module-subtitle">Scoop / retro</span>
                <span className="module-ring"><strong>{recordCounts.rezagado}</strong><small>reg.</small></span>
              </button>
              <button
                type="button"
                className={`module-card module-yellow ${formType === 'seguridad' ? 'active' : ''}`}
                onClick={() => selectFormType('seguridad')}
              >
                <span className="module-icon"><ShieldCheck size={24} /></span>
                <span className="module-title">Seguridad</span>
                <span className="module-subtitle">Incidentes</span>
                <span className="module-ring"><strong>{recordCounts.seguridad}</strong><small>reg.</small></span>
              </button>
              <button
                type="button"
                className="module-card module-green"
                onClick={() => {
                  setSection('dashboard')
                  void syncRecords({ silent: true })
                }}
              >
                <span className="module-icon"><Pickaxe size={24} /></span>
                <span className="module-title">Revision</span>
                <span className="module-subtitle">KPI / reportes</span>
                <span className="module-ring"><strong>{recordCounts.total}</strong><small>total</small></span>
              </button>
            </div>
            <button className="primary-action" type="submit">
              <Save size={18} /> {editingId ? 'Guardar cambios' : 'Guardar registro'}
            </button>
            {editingId && (
              <button className="secondary-action" type="button" onClick={cancelEdit}>
                <X size={18} /> Cancelar edicion
              </button>
            )}
            <div className="recent-captures">
              <div className="recent-title">
                <span>Capturas recientes</span>
                <strong>{recordCounts.total}</strong>
              </div>
              {recentRecords.length === 0 ? (
                <p>Sin capturas todavia.</p>
              ) : (
                recentRecords.map((record) => (
                  <article className="recent-card" key={record.id}>
                    <button type="button" onClick={() => startEdit(record)}>
                      <span>{labelType(record.type)}</span>
                      <strong>{record.fecha} · Turno {record.turno}</strong>
                    </button>
                    <button className="icon-button danger" type="button" onClick={() => removeRecord(record.id)} aria-label="Eliminar captura">
                      <Trash2 size={16} />
                    </button>
                  </article>
                ))
              )}
            </div>
          </aside>

          <section className="form-panel">
            {formType === 'barrenacion' && (
              <BarrenacionForm record={barrenacion} setRecord={setBarrenacion} />
            )}
            {formType === 'rezagado' && <RezagadoForm record={rezagado} setRecord={setRezagado} />}
            {formType === 'seguridad' && <SeguridadForm record={seguridad} setRecord={setSeguridad} />}
          </section>
        </form>
      ) : (
        <section className="dashboard">
          <div className="dashboard-tools">
            <div>
              <h1>Revision y KPI</h1>
              <p>Panel local o web Render para consolidar registros sincronizados.</p>
            </div>
            <div className="tool-actions">
              <input
                aria-label="Buscar registros"
                placeholder="Buscar por fecha, supervisor, turno..."
                value={query}
                onChange={(event) => setQuery(event.target.value)}
              />
              <input
                aria-label="URL Render"
                placeholder="URL Render"
                value={apiUrl}
                onChange={(event) => setApiUrl(event.target.value)}
              />
              <button onClick={() => void syncRecords()} disabled={syncing}>
                <RefreshCw size={18} className={syncing ? 'spin' : ''} /> Sincronizar
              </button>
              <button onClick={exportJson}>
                <Download size={18} /> JSON
              </button>
              <label className="file-button">
                <Upload size={18} /> Importar
                <input type="file" accept="application/json" onChange={importJson} />
              </label>
              <button onClick={exportPdf}>
                <FileDown size={18} /> PDF
              </button>
            </div>
          </div>

          <div ref={reportRef} className="report-page">
            <div className="report-header">
              <img src="/mga-logo.jfif" alt="MGA" />
              <div>
                <span>MGA CONTRATISTA MINERA S.A. DE C.V.</span>
                <strong>Reporte ejecutivo de operaciones mina</strong>
              </div>
              <small>{new Date().toLocaleDateString('es-MX')}</small>
            </div>

            <div className="kpi-grid">
              <Kpi icon={<History />} label="Registros" value={kpis.total} />
              <Kpi icon={<BarChart3 />} label="Metros barrenados" value={kpis.metrosDados} />
              <Kpi icon={<Activity />} label="Metros pegados" value={kpis.metrosPegados} />
              <Kpi icon={<HardHat />} label="Cucharones rezagado" value={kpis.rezagado} />
              <Kpi icon={<ShieldCheck />} label="Accidentes / Incidentes" value={`${kpis.accidentes} / ${kpis.incidentes}`} />
              <Kpi icon={<ClipboardCheck />} label="Fuerza laboral" value={kpis.fuerzaLaboral} />
            </div>

            <div className="summary-grid">
              <section>
                <h2>Produccion</h2>
                <Metric label="Barrenos dados" value={kpis.barrenosDados} />
                <Metric label="Barrenos cargados" value={kpis.barrenosCargados} />
                <Metric label="Camiones" value={kpis.camiones} />
                <Metric label="Diesel" value={kpis.diesel} />
              </section>
              <section>
                <h2>Seguridad</h2>
                <Metric label="Accidentes" value={kpis.accidentes} />
                <Metric label="Incidentes" value={kpis.incidentes} />
                <Metric label="Reportes de seguridad" value={kpis.seguridad} />
                <Metric label="Personas registradas" value={kpis.fuerzaLaboral} />
              </section>
            </div>

            <div className="records-table">
              <h2>Registros capturados</h2>
              <table>
                <thead>
                  <tr>
                    <th>Fecha</th>
                    <th>Turno</th>
                    <th>Tipo</th>
                    <th>Supervisor</th>
                    <th>Unidad</th>
                    <th>Estado</th>
                    <th data-html2canvas-ignore="true">Acciones</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((record) => (
                    <tr key={record.id}>
                      <td>{record.fecha}</td>
                      <td>{record.turno}</td>
                      <td>{labelType(record.type)}</td>
                      <td>{record.supervisor || 'Pendiente'}</td>
                      <td>{record.unidad}</td>
                      <td>
                        <span className={record.updatedAt === record.syncedAt ? 'status-pill synced' : 'status-pill'}>
                          {record.updatedAt === record.syncedAt ? 'En red' : 'Pendiente'}
                        </span>
                      </td>
                      <td className="table-actions" data-html2canvas-ignore="true">
                        <button className="icon-button" onClick={() => startEdit(record)} aria-label="Editar">
                          <Edit3 size={16} />
                        </button>
                        <button className="icon-button danger" onClick={() => removeRecord(record.id)} aria-label="Eliminar">
                          <Trash2 size={16} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </section>
      )}
    </main>
  )
}

function BaseFields<T extends BaseRecord>({
  record,
  setRecord,
}: {
  record: T
  setRecord: (record: T) => void
}) {
  return (
    <div className="form-grid four">
      <Field label="Supervisor" value={record.supervisor} onChange={(value) => setRecord({ ...record, supervisor: value })} required />
      <Field label="Fecha" type="date" value={record.fecha} onChange={(value) => setRecord({ ...record, fecha: value })} required />
      <label>
        Turno
        <select value={record.turno} onChange={(event) => setRecord({ ...record, turno: event.target.value as Shift })}>
          <option value="1">Turno 1</option>
          <option value="2">Turno 2</option>
        </select>
      </label>
      <Field label="Unidad" value={record.unidad} onChange={(value) => setRecord({ ...record, unidad: value })} />
    </div>
  )
}

function BarrenacionForm({ record, setRecord }: { record: BarrenacionRecord; setRecord: (record: BarrenacionRecord) => void }) {
  const activity = getBarrenacionActivity(record)
  const drillRow =
    activity === 'maquinaPierna'
      ? record.maquinaPierna[0] ?? emptyDrillRow(defaultJumbo)
      : record.jumbo[0] ?? emptyDrillRow(defaultJumbo)
  const blastRow = record.voladuras[0] ?? emptyBlastRow()

  function selectActivity(next: BarrenacionActivity) {
    if (next === 'voladura') {
      setRecord({ ...record, jumbo: [], maquinaPierna: [], voladuras: [record.voladuras[0] ?? emptyBlastRow()] })
      return
    }
    const currentDrill = (next === 'maquinaPierna' ? record.maquinaPierna[0] : record.jumbo[0])
      ?? record.jumbo[0]
      ?? record.maquinaPierna[0]
      ?? emptyDrillRow(defaultJumbo)
    setRecord({
      ...record,
      jumbo: next === 'jumbo' ? [currentDrill] : [],
      maquinaPierna: next === 'maquinaPierna' ? [currentDrill] : [],
      voladuras: [],
    })
  }

  function updateDrill(patch: Partial<DrillRow>) {
    const updated = { ...drillRow, ...patch }
    setRecord({
      ...record,
      jumbo: activity === 'jumbo' ? [updated] : record.jumbo,
      maquinaPierna: activity === 'maquinaPierna' ? [updated] : record.maquinaPierna,
    })
  }

  function updateBlast(patch: Partial<BlastRow>) {
    setRecord({ ...record, voladuras: [{ ...blastRow, ...patch }] })
  }

  return (
    <>
      <PanelTitle title="Captura rapida de barrenos" subtitle="Evento corto de campo" />
      <QuickSection title="Turno" icon={<Mountain size={18} />}>
        <BaseFields record={record} setRecord={setRecord} />
      </QuickSection>
      <QuickSection title="Actividad" icon={<Drill size={18} />}>
        <div className="quick-choice-grid three">
          <ChoiceButton active={activity === 'jumbo'} icon={<Drill size={21} />} label="Jumbo" meta="Barrenacion" onClick={() => selectActivity('jumbo')} />
          <ChoiceButton active={activity === 'maquinaPierna'} icon={<HardHat size={21} />} label="Maquina pierna" meta="Barrenacion" onClick={() => selectActivity('maquinaPierna')} />
          <ChoiceButton active={activity === 'voladura'} icon={<Activity size={21} />} label="Voladura" meta="Pegada" onClick={() => selectActivity('voladura')} />
        </div>
      </QuickSection>

      {activity === 'voladura' ? (
        <QuickSection title="Produccion" icon={<Activity size={18} />}>
          <div className="form-grid quick">
            <Field label="Obra / frente" value={blastRow.obra} onChange={(value) => updateBlast({ obra: value })} required />
            <Field label="Oficial" value={blastRow.oficial} onChange={(value) => updateBlast({ oficial: value })} />
            <Field label="Ayudante" value={blastRow.ayudante} onChange={(value) => updateBlast({ ayudante: value })} />
            <Field label="RPA / Cfte" value={blastRow.rpaCfte} onChange={(value) => updateBlast({ rpaCfte: value })} />
            <Field label="Barrenos pegados" type="number" value={blastRow.barrenosPegados} onChange={(value) => updateBlast({ barrenosPegados: Number(value) })} />
            <Field label="Metros pegados" type="number" value={blastRow.metrosPegados} onChange={(value) => updateBlast({ metrosPegados: Number(value) })} />
            <Field label="Horas servicio" type="number" value={blastRow.horasServicio} onChange={(value) => updateBlast({ horasServicio: Number(value) })} />
          </div>
          <details className="quick-details">
            <summary>Insumos y avance</summary>
            <div className="form-grid quick">
              <Field label="Longitud" type="number" value={blastRow.longitud} onChange={(value) => updateBlast({ longitud: Number(value) })} />
              <Field label="Cuele" type="number" value={blastRow.cuele} onChange={(value) => updateBlast({ cuele: Number(value) })} />
              <Field label="Desarrollo" type="number" value={blastRow.desarrollo} onChange={(value) => updateBlast({ desarrollo: Number(value) })} />
              <Field label="ANFO inicial" type="number" value={blastRow.anfoInicial} onChange={(value) => updateBlast({ anfoInicial: Number(value) })} />
              <Field label="ANFO final" type="number" value={blastRow.anfoFinal} onChange={(value) => updateBlast({ anfoFinal: Number(value) })} />
            </div>
          </details>
          <TextArea label="Observaciones" value={blastRow.observaciones} onChange={(value) => updateBlast({ observaciones: value })} />
        </QuickSection>
      ) : (
        <QuickSection title="Produccion" icon={<Pickaxe size={18} />}>
          <div className="form-grid quick">
            <Field label="Equipo" value={drillRow.equipo} options={jumboEquipoOptions} onChange={(value) => updateDrill({ equipo: value })} required />
            <Field label="Operador" value={drillRow.operador} onChange={(value) => updateDrill({ operador: value })} required />
            <Field label="Ayudante" value={drillRow.ayudante} onChange={(value) => updateDrill({ ayudante: value })} />
            <Field label="Nivel / obra" value={drillRow.nivelObra} onChange={(value) => updateDrill({ nivelObra: value })} required />
            <Field label="RPA / Cfte" value={drillRow.rpaCfte} onChange={(value) => updateDrill({ rpaCfte: value })} />
            <Field label="Barrenos dados" type="number" value={drillRow.barrenosDados} onChange={(value) => updateDrill({ barrenosDados: Number(value) })} />
            <Field label="Barrenos cargados" type="number" value={drillRow.barrenosCargados} onChange={(value) => updateDrill({ barrenosCargados: Number(value) })} />
            <Field label="Metros dados" type="number" value={drillRow.metrosDados} onChange={(value) => updateDrill({ metrosDados: Number(value) })} />
            <Field label="Horas servicio" type="number" value={drillRow.horasServicio} onChange={(value) => updateDrill({ horasServicio: Number(value) })} />
          </div>
          <details className="quick-details">
            <summary>Horometros e insumos</summary>
            <div className="form-grid quick">
              <Field label="Longitud" type="number" value={drillRow.longitud} onChange={(value) => updateDrill({ longitud: Number(value) })} />
              <Field label="Cuele" type="number" value={drillRow.cuele} onChange={(value) => updateDrill({ cuele: Number(value) })} />
              <Field label="Desarrollo" type="number" value={drillRow.desarrollo} onChange={(value) => updateDrill({ desarrollo: Number(value) })} />
              <Field label="Hor. diesel inicial" type="number" value={drillRow.horometroDieselInicial} onChange={(value) => updateDrill({ horometroDieselInicial: Number(value) })} />
              <Field label="Hor. diesel final" type="number" value={drillRow.horometroDieselFinal} onChange={(value) => updateDrill({ horometroDieselFinal: Number(value) })} />
              <Field label="Zanco" type="number" value={drillRow.zanco} onChange={(value) => updateDrill({ zanco: Number(value) })} />
              <Field label="Cople" type="number" value={drillRow.cople} onChange={(value) => updateDrill({ cople: Number(value) })} />
              <Field label="Barra" type="number" value={drillRow.barra} onChange={(value) => updateDrill({ barra: Number(value) })} />
              <Field label="Broca" type="number" value={drillRow.broca} onChange={(value) => updateDrill({ broca: Number(value) })} />
            </div>
          </details>
          <TextArea label="Comentarios" value={record.comentarios} onChange={(value) => setRecord({ ...record, comentarios: value })} />
        </QuickSection>
      )}
    </>
  )
}

function RezagadoForm({ record, setRecord }: { record: RezagadoRecord; setRecord: (record: RezagadoRecord) => void }) {
  const equipment = getRezagadoEquipment(record)
  const scoopRow = record.scoopTram[0] ?? emptyHaulRow(defaultScoop)
  const retroRow = record.retro[0] ?? emptyRetroRow()
  const haulActivity = getActiveKey(scoopRow, haulActivityKeys, 'rezagado')
  const retroActivity = getActiveKey(retroRow, retroActivityKeys, 'amacice')

  function selectEquipment(next: RezagadoEquipment) {
    setRecord({
      ...record,
      scoopTram: next === 'scoop' ? [record.scoopTram[0] ?? emptyHaulRow(defaultScoop)] : [],
      retro: next === 'retro' ? [record.retro[0] ?? emptyRetroRow()] : [],
    })
  }

  function updateScoop(row: HaulRow) {
    setRecord({ ...record, scoopTram: [row] })
  }

  function updateRetro(row: RetroRow) {
    setRecord({ ...record, retro: [row] })
  }

  return (
    <>
      <PanelTitle title="Captura rapida de rezagado" subtitle="Equipo, ubicacion y produccion" />
      <QuickSection title="Turno" icon={<Mountain size={18} />}>
        <BaseFields record={record} setRecord={setRecord} />
      </QuickSection>
      <QuickSection title="Equipo" icon={<Truck size={18} />}>
        <div className="quick-choice-grid two">
          <ChoiceButton active={equipment === 'scoop'} icon={<Truck size={21} />} label="Scoop tram" meta="Rezagado" onClick={() => selectEquipment('scoop')} />
          <ChoiceButton active={equipment === 'retro'} icon={<HardHat size={21} />} label="Retro" meta="Obra civil" onClick={() => selectEquipment('retro')} />
        </div>
        {equipment === 'scoop' ? (
          <div className="form-grid quick">
            <Field label="Equipo" value={scoopRow.equipo} options={scoopEquipoOptions} onChange={(value) => updateScoop({ ...scoopRow, equipo: value })} required />
            <Field label="Operador" value={scoopRow.operador} onChange={(value) => updateScoop({ ...scoopRow, operador: value })} required />
            <Field label="Nivel / obra" value={scoopRow.nivelObra} onChange={(value) => updateScoop({ ...scoopRow, nivelObra: value })} required />
            <Field label="Destino" value={scoopRow.destino} onChange={(value) => updateScoop({ ...scoopRow, destino: value })} />
          </div>
        ) : (
          <div className="form-grid quick">
            <Field label="Equipo" value={retroRow.equipo} options={retroEquipoOptions} onChange={(value) => updateRetro({ ...retroRow, equipo: value })} required />
            <Field label="Operador" value={retroRow.operador} onChange={(value) => updateRetro({ ...retroRow, operador: value })} required />
            <Field label="Nivel / obra" value={retroRow.nivelObra} onChange={(value) => updateRetro({ ...retroRow, nivelObra: value })} required />
          </div>
        )}
      </QuickSection>
      <QuickSection title="Trabajo" icon={<Activity size={18} />}>
        {equipment === 'scoop' ? (
          <>
            <div className="quick-choice-grid compact">
              {haulActivityKeys.map((key) => (
                <ChoiceButton
                  key={key}
                  active={haulActivity === key}
                  label={haulActivityLabels[key]}
                  onClick={() => updateScoop(withSingleMetric(scoopRow, haulActivityKeys, key, Number(scoopRow[haulActivity] || 0)))}
                />
              ))}
            </div>
            <div className="form-grid quick">
              <Field label="Cantidad" type="number" value={scoopRow[haulActivity]} onChange={(value) => updateScoop(withSingleMetric(scoopRow, haulActivityKeys, haulActivity, Number(value)))} />
              <Field label="Camiones" type="number" value={scoopRow.camiones} onChange={(value) => updateScoop({ ...scoopRow, camiones: Number(value) })} />
              <Field label="Hor. inicial" type="number" value={scoopRow.horometroInicial} onChange={(value) => updateScoop({ ...scoopRow, horometroInicial: Number(value) })} />
              <Field label="Hor. final" type="number" value={scoopRow.horometroFinal} onChange={(value) => updateScoop({ ...scoopRow, horometroFinal: Number(value) })} />
              <Field label="Diesel" type="number" value={scoopRow.diesel} onChange={(value) => updateScoop({ ...scoopRow, diesel: Number(value) })} />
            </div>
            <TextArea label="Observaciones" value={scoopRow.observaciones} onChange={(value) => updateScoop({ ...scoopRow, observaciones: value })} />
          </>
        ) : (
          <>
            <div className="quick-choice-grid compact">
              {retroActivityKeys.map((key) => (
                <ChoiceButton
                  key={key}
                  active={retroActivity === key}
                  label={retroActivityLabels[key]}
                  onClick={() => updateRetro(withSingleMetric(retroRow, retroActivityKeys, key, Number(retroRow[retroActivity] || 0)))}
                />
              ))}
            </div>
            <div className="form-grid quick">
              <Field label="Cantidad" type="number" value={retroRow[retroActivity]} onChange={(value) => updateRetro(withSingleMetric(retroRow, retroActivityKeys, retroActivity, Number(value)))} />
              <Field label="Hor. inicial" type="number" value={retroRow.horometroInicial} onChange={(value) => updateRetro({ ...retroRow, horometroInicial: Number(value) })} />
              <Field label="Hor. final" type="number" value={retroRow.horometroFinal} onChange={(value) => updateRetro({ ...retroRow, horometroFinal: Number(value) })} />
              <Field label="Diesel" type="number" value={retroRow.diesel} onChange={(value) => updateRetro({ ...retroRow, diesel: Number(value) })} />
            </div>
            <TextArea label="Observaciones" value={retroRow.observaciones} onChange={(value) => updateRetro({ ...retroRow, observaciones: value })} />
          </>
        )}
      </QuickSection>
      <TextArea label="Comentarios generales" value={record.comentarios} onChange={(value) => setRecord({ ...record, comentarios: value })} />
    </>
  )
}

function SeguridadForm({ record, setRecord }: { record: SeguridadRecord; setRecord: (record: SeguridadRecord) => void }) {
  return (
    <>
      <PanelTitle title="Captura rapida de seguridad" subtitle="Eventos y acciones preventivas" />
      <QuickSection title="Turno" icon={<Mountain size={18} />}>
        <BaseFields record={record} setRecord={setRecord} />
      </QuickSection>
      <QuickSection title="Resumen" icon={<ShieldCheck size={18} />}>
        <div className="form-grid quick">
          <Field label="Accidentes" type="number" value={record.accidentes} onChange={(value) => setRecord({ ...record, accidentes: Number(value) })} />
          <Field label="Incidentes" type="number" value={record.incidentes} onChange={(value) => setRecord({ ...record, incidentes: Number(value) })} />
          <Field label="Fuerza laboral" type="number" value={record.fuerzaLaboral} onChange={(value) => setRecord({ ...record, fuerzaLaboral: Number(value) })} />
        </div>
      </QuickSection>
      <QuickSection title="Hallazgo" icon={<ClipboardCheck size={18} />}>
        <TextArea label="Acto o condicion insegura" value={record.actosInseguros} onChange={(value) => setRecord({ ...record, actosInseguros: value })} />
        <TextArea label="Accion correctiva" value={record.correccionesMejoras} onChange={(value) => setRecord({ ...record, correccionesMejoras: value })} />
        <details className="quick-details">
          <summary>Actividades y platica</summary>
          <TextArea label="Platica de seguridad" value={record.platicaSeguridad} onChange={(value) => setRecord({ ...record, platicaSeguridad: value })} />
          <TextArea label="Actividades de seguridad" value={record.actividadesSeguridad} onChange={(value) => setRecord({ ...record, actividadesSeguridad: value })} />
          <TextArea label="Actividades de operacion" value={record.actividadesOperacion} onChange={(value) => setRecord({ ...record, actividadesOperacion: value })} />
        </details>
        <TextArea label="Observaciones" value={record.observaciones} onChange={(value) => setRecord({ ...record, observaciones: value })} />
      </QuickSection>
    </>
  )
}

function QuickSection({ title, icon, children }: { title: string; icon: ReactNode; children: ReactNode }) {
  return (
    <section className="quick-section">
      <div className="quick-section-title">
        <span>{icon}</span>
        <h2>{title}</h2>
      </div>
      {children}
    </section>
  )
}

function ChoiceButton({
  active,
  icon,
  label,
  meta,
  onClick,
}: {
  active: boolean
  icon?: ReactNode
  label: string
  meta?: string
  onClick: () => void
}) {
  return (
    <button type="button" className={`choice-button ${active ? 'active' : ''}`} onClick={onClick}>
      {icon && <span>{icon}</span>}
      <strong>{label}</strong>
      {meta && <small>{meta}</small>}
    </button>
  )
}

function Field({
  label,
  value,
  onChange,
  type = 'text',
  required = false,
  options,
}: {
  label: string
  value: string | number
  onChange: (value: string) => void
  type?: string
  required?: boolean
  options?: string[]
}) {
  const textValue = String(value)
  return (
    <label>
      {label}
      {options ? (
        <select className="scroll-select" required={required} value={textValue} onChange={(event) => onChange(event.target.value)}>
          <option value="">Seleccionar equipo</option>
          {textValue && !options.includes(textValue) && <option value={textValue}>{textValue}</option>}
          {options.map((option) => (
            <option key={option} value={option}>
              {option}
            </option>
          ))}
        </select>
      ) : (
        <input
          required={required}
          type={type}
          value={value}
          min={type === 'number' ? 0 : undefined}
          step={type === 'number' ? '0.01' : undefined}
          onChange={(event) => onChange(event.target.value)}
        />
      )}
    </label>
  )
}

function TextArea({ label, value, onChange }: { label: string; value: string; onChange: (value: string) => void }) {
  return (
    <label className="textarea-field">
      {label}
      <textarea value={value} onChange={(event) => onChange(event.target.value)} rows={4} />
    </label>
  )
}

function PanelTitle({ title, subtitle }: { title: string; subtitle: string }) {
  return (
    <div className="panel-title">
      <span>{subtitle}</span>
      <h1>{title}</h1>
    </div>
  )
}

function Kpi({ icon, label, value }: { icon: ReactNode; label: string; value: string | number }) {
  return (
    <div className="kpi-card">
      {icon}
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  )
}

function Metric({ label, value }: { label: string; value: number }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  )
}

function loadRecords(): MineRecord[] {
  try {
    return (JSON.parse(localStorage.getItem(STORAGE_KEY) ?? '[]') as MineRecord[]).map(normalizeRecord)
  } catch {
    return []
  }
}

function getInitialSection(): 'captura' | 'dashboard' {
  const params = new URLSearchParams(window.location.search)
  const requestedView = params.get('vista')
  if (requestedView === 'captura') return 'captura'
  if (requestedView === 'revision' || requestedView === 'dashboard') return 'dashboard'
  return Capacitor.isNativePlatform() ? 'captura' : 'dashboard'
}

function loadApiUrl() {
  return localStorage.getItem(API_URL_KEY) ?? import.meta.env.VITE_API_URL ?? DEFAULT_API_URL
}

function resolveApiBase(apiUrl: string) {
  return (apiUrl.trim() || DEFAULT_API_URL).replace(/\/$/, '')
}

function normalizeRecord(record: MineRecord): MineRecord {
  const createdAt = record.createdAt ?? nowIso()
  return {
    ...record,
    createdAt,
    updatedAt: record.updatedAt ?? createdAt,
  } as MineRecord
}

function mergeRecords(records: MineRecord[]) {
  const map = new Map<string, MineRecord>()
  records.map(normalizeRecord).forEach((record) => {
    const current = map.get(record.id)
    if (!current || new Date(record.updatedAt).getTime() >= new Date(current.updatedAt).getTime()) {
      map.set(record.id, record)
    }
  })
  return Array.from(map.values()).sort((a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime())
}

function sum<T>(rows: T[], key: keyof T) {
  return rows.reduce((total, row) => total + Number(row[key] || 0), 0)
}

function computeKpis(records: MineRecord[]) {
  return records.reduce(
    (acc, record) => {
      acc.total += 1
      if (record.type === 'barrenacion') {
        const drillRows = [...record.jumbo, ...record.maquinaPierna]
        acc.metrosDados += sum(drillRows, 'metrosDados')
        acc.barrenosDados += sum(drillRows, 'barrenosDados')
        acc.barrenosCargados += sum(drillRows, 'barrenosCargados')
        acc.metrosPegados += sum(record.voladuras, 'metrosPegados')
      }
      if (record.type === 'rezagado') {
        acc.rezagado += sum(record.scoopTram, 'rezagado')
        acc.camiones += sum(record.scoopTram, 'camiones')
        acc.diesel += sum(record.scoopTram, 'diesel') + sum(record.retro, 'diesel')
      }
      if (record.type === 'seguridad') {
        acc.seguridad += 1
        acc.accidentes += Number(record.accidentes || 0)
        acc.incidentes += Number(record.incidentes || 0)
        acc.fuerzaLaboral += Number(record.fuerzaLaboral || 0)
      }
      return acc
    },
    {
      total: 0,
      metrosDados: 0,
      metrosPegados: 0,
      barrenosDados: 0,
      barrenosCargados: 0,
      rezagado: 0,
      camiones: 0,
      diesel: 0,
      seguridad: 0,
      accidentes: 0,
      incidentes: 0,
      fuerzaLaboral: 0,
    },
  )
}

function labelType(type: RecordType) {
  return {
    barrenacion: 'Barrenacion y voladuras',
    rezagado: 'Rezagado retro',
    seguridad: 'Seguridad',
  }[type]
}

const haulActivityKeys = ['rezagado', 'traspaleo', 'limpia', 'balastreo', 'planilla', 'relleno'] as const
const retroActivityKeys = ['amacice', 'reAmacice', 'tableo', 'limpia', 'balastreo', 'mtAcequia'] as const

const haulActivityLabels: Record<HaulActivity, string> = {
  rezagado: 'Rezagado',
  traspaleo: 'Traspaleo',
  limpia: 'Limpia',
  balastreo: 'Balastreo',
  planilla: 'Planilla',
  relleno: 'Relleno',
}

const retroActivityLabels: Record<RetroActivity, string> = {
  amacice: 'Amacice',
  reAmacice: 'Re-amacice',
  tableo: 'Tableo',
  limpia: 'Limpia',
  balastreo: 'Balastreo',
  mtAcequia: 'Mt acequia',
}

function hasDrillData(row: DrillRow) {
  return Boolean(
    row.operador
      || row.ayudante
      || row.nivelObra
      || row.rpaCfte
      || row.barrenosDados
      || row.barrenosCargados
      || row.metrosDados
      || row.horasServicio,
  )
}

function hasBlastData(row: BlastRow) {
  return Boolean(
    row.obra
      || row.oficial
      || row.ayudante
      || row.rpaCfte
      || row.barrenosPegados
      || row.metrosPegados
      || row.horasServicio
      || row.observaciones,
  )
}

function hasHaulData(row: HaulRow) {
  return Boolean(
    row.operador
      || row.nivelObra
      || row.destino
      || row.rezagado
      || row.traspaleo
      || row.limpia
      || row.balastreo
      || row.planilla
      || row.relleno
      || row.camiones
      || row.diesel,
  )
}

function hasRetroData(row: RetroRow) {
  return Boolean(
    row.operador
      || row.nivelObra
      || row.amacice
      || row.reAmacice
      || row.tableo
      || row.limpia
      || row.balastreo
      || row.mtAcequia
      || row.diesel,
  )
}

function getBarrenacionActivity(record: BarrenacionRecord): BarrenacionActivity {
  if (record.voladuras.some(hasBlastData)) return 'voladura'
  if (record.maquinaPierna.some(hasDrillData)) return 'maquinaPierna'
  return 'jumbo'
}

function getRezagadoEquipment(record: RezagadoRecord): RezagadoEquipment {
  if (record.retro.some(hasRetroData) && !record.scoopTram.some(hasHaulData)) return 'retro'
  return 'scoop'
}

function getActiveKey<T extends Record<string, string | number>, K extends keyof T>(
  row: T,
  keys: readonly K[],
  fallback: K,
) {
  return keys.find((key) => Number(row[key] || 0) > 0) ?? fallback
}

function withSingleMetric<T extends Record<string, string | number>, K extends keyof T>(
  row: T,
  keys: readonly K[],
  activeKey: K,
  value: number,
) {
  return keys.reduce(
    (next, key) => ({
      ...next,
      [key]: key === activeKey ? value : 0,
    }),
    { ...row },
  )
}

export default App
