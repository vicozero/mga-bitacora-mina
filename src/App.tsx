import { useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent, FormEvent, ReactNode } from 'react'
import {
  Activity,
  BarChart3,
  ClipboardCheck,
  Download,
  Edit3,
  FileDown,
  HardHat,
  History,
  LayoutDashboard,
  Plus,
  RefreshCw,
  Save,
  ShieldCheck,
  Trash2,
  Upload,
  X,
} from 'lucide-react'
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
type RowField<T> = { key: keyof T; label: string; type?: string; options?: string[] }

const STORAGE_KEY = 'mga-bitacora-operaciones-v1'
const API_URL_KEY = 'mga-bitacora-api-url'
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
    maquinaPierna: [emptyDrillRow(defaultJumbo)],
    voladuras: [emptyBlastRow()],
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
    retro: [emptyRetroRow()],
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
  const [section, setSection] = useState<'captura' | 'dashboard'>('captura')
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
  const pendingSync = useMemo(
    () => records.filter((record) => record.updatedAt !== record.syncedAt).length,
    [records],
  )

  useEffect(() => {
    localStorage.setItem(API_URL_KEY, apiUrl)
  }, [apiUrl])

  function persist(next: MineRecord[]) {
    const normalized = mergeRecords(next.map(normalizeRecord))
    setRecords(normalized)
    localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized))
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
    persist(next)
    resetForm(formType)
    setEditingId(null)
    setMessage(editingId ? 'Captura actualizada offline.' : 'Registro guardado offline en este dispositivo.')
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
    persist(
      records.map((record) =>
        record.id === id
          ? normalizeRecord({ ...record, deletedAt, updatedAt: deletedAt, syncedAt: undefined } as MineRecord)
          : record,
      ),
    )
    setMessage('Captura eliminada localmente. Sincroniza para eliminarla en Render.')
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
    persist([...imported.map(normalizeRecord), ...records])
    setMessage('Datos importados al panel local.')
    event.target.value = ''
  }

  async function syncRecords() {
    const apiBase = resolveApiBase(apiUrl)
    if (window.location.protocol === 'capacitor:' && !apiBase) {
      setMessage('Configura la URL de Render antes de sincronizar desde la APK.')
      return
    }
    setSyncing(true)
    try {
      const localRecords = records.map(normalizeRecord)
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
      setMessage('Sincronizacion completada con Render.')
    } catch {
      setMessage('No se pudo sincronizar. Revisa internet y la URL de Render.')
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
    <main className="app-shell">
      <header className="topbar">
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
          <button className={section === 'dashboard' ? 'active' : ''} onClick={() => setSection('dashboard')}>
            <LayoutDashboard size={18} /> Revision web
          </button>
        </nav>
        <button className="sync-button" onClick={syncRecords} disabled={syncing}>
          <RefreshCw size={18} className={syncing ? 'spin' : ''} /> Sincronizar
          {pendingSync > 0 && <span>{pendingSync}</span>}
        </button>
      </header>

      {message && <div className="toast">{message}</div>}

      {section === 'captura' ? (
        <form className="workspace" onSubmit={saveRecord}>
          <aside className="side-panel">
            <h1>{editingId ? 'Editar captura' : 'Captura offline'}</h1>
            <p>Guarda sin internet y sincroniza con Render cuando haya red.</p>
            <label className="server-field">
              URL Render
              <input
                placeholder="https://tu-servicio.onrender.com"
                value={apiUrl}
                onChange={(event) => setApiUrl(event.target.value)}
              />
            </label>
            <div className="record-tabs">
              <button type="button" className={formType === 'barrenacion' ? 'active' : ''} onClick={() => selectFormType('barrenacion')}>
                <HardHat size={18} /> Barrenacion y voladuras
              </button>
              <button type="button" className={formType === 'rezagado' ? 'active' : ''} onClick={() => selectFormType('rezagado')}>
                <Activity size={18} /> Rezagado retro
              </button>
              <button type="button" className={formType === 'seguridad' ? 'active' : ''} onClick={() => selectFormType('seguridad')}>
                <ShieldCheck size={18} /> Seguridad
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
              <button onClick={syncRecords} disabled={syncing}>
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
  return (
    <>
      <PanelTitle title="Reporte Diario de Barrenacion y Voladuras" subtitle="Basado en el formato original MGA" />
      <BaseFields record={record} setRecord={setRecord} />
      <EditableRows
        title="Jumbo"
        rows={record.jumbo}
        onAdd={() => setRecord({ ...record, jumbo: [...record.jumbo, emptyDrillRow(defaultJumbo)] })}
        onChange={(jumbo) => setRecord({ ...record, jumbo })}
        fields={drillFields(jumboEquipoOptions)}
      />
      <EditableRows
        title="Maquina pierna"
        rows={record.maquinaPierna}
        onAdd={() => setRecord({ ...record, maquinaPierna: [...record.maquinaPierna, emptyDrillRow(defaultJumbo)] })}
        onChange={(maquinaPierna) => setRecord({ ...record, maquinaPierna })}
        fields={drillFields(jumboEquipoOptions)}
      />
      <EditableRows
        title="Voladuras"
        rows={record.voladuras}
        onAdd={() => setRecord({ ...record, voladuras: [...record.voladuras, emptyBlastRow()] })}
        onChange={(voladuras) => setRecord({ ...record, voladuras })}
        fields={blastFields}
      />
      <div className="form-grid three">
        {(['polvorero', 'choferCamion', 'choferPipa', 'bobCat', 'bombeo', 'servicios'] as const).map((key) => (
          <Field key={key} label={activityLabels[key]} value={record[key]} onChange={(value) => setRecord({ ...record, [key]: value })} />
        ))}
      </div>
      <TextArea label="Contratiempos de operacion / comentarios generales" value={record.comentarios} onChange={(value) => setRecord({ ...record, comentarios: value })} />
      <TextArea label="Inasistencias / permisos" value={record.inasistencias} onChange={(value) => setRecord({ ...record, inasistencias: value })} />
    </>
  )
}

function RezagadoForm({ record, setRecord }: { record: RezagadoRecord; setRecord: (record: RezagadoRecord) => void }) {
  return (
    <>
      <PanelTitle title="Reporte Diario de Rezagado" subtitle="Scoop tram, retro, horometros y combustible" />
      <BaseFields record={record} setRecord={setRecord} />
      <EditableRows
        title="Scoop tram"
        rows={record.scoopTram}
        onAdd={() => setRecord({ ...record, scoopTram: [...record.scoopTram, emptyHaulRow(defaultScoop)] })}
        onChange={(scoopTram) => setRecord({ ...record, scoopTram })}
        fields={haulFields}
      />
      <EditableRows
        title="Retro"
        rows={record.retro}
        onAdd={() => setRecord({ ...record, retro: [...record.retro, emptyRetroRow()] })}
        onChange={(retro) => setRecord({ ...record, retro })}
        fields={retroFields}
      />
      <TextArea label="Observaciones y comentarios" value={record.comentarios} onChange={(value) => setRecord({ ...record, comentarios: value })} />
    </>
  )
}

function SeguridadForm({ record, setRecord }: { record: SeguridadRecord; setRecord: (record: SeguridadRecord) => void }) {
  return (
    <>
      <PanelTitle title="Reporte de Seguridad" subtitle="Eventos, fuerza laboral y actividades preventivas" />
      <BaseFields record={record} setRecord={setRecord} />
      <div className="form-grid three">
        <Field label="Accidentes" type="number" value={record.accidentes} onChange={(value) => setRecord({ ...record, accidentes: Number(value) })} />
        <Field label="Incidentes" type="number" value={record.incidentes} onChange={(value) => setRecord({ ...record, incidentes: Number(value) })} />
        <Field label="Fuerza laboral" type="number" value={record.fuerzaLaboral} onChange={(value) => setRecord({ ...record, fuerzaLaboral: Number(value) })} />
      </div>
      <TextArea label="Actos inseguros" value={record.actosInseguros} onChange={(value) => setRecord({ ...record, actosInseguros: value })} />
      <TextArea label="Condiciones inseguras" value={record.condicionesInseguras} onChange={(value) => setRecord({ ...record, condicionesInseguras: value })} />
      <TextArea label="Platica de seguridad" value={record.platicaSeguridad} onChange={(value) => setRecord({ ...record, platicaSeguridad: value })} />
      <TextArea label="Actividades de seguridad" value={record.actividadesSeguridad} onChange={(value) => setRecord({ ...record, actividadesSeguridad: value })} />
      <TextArea label="Correcciones y/o mejoras" value={record.correccionesMejoras} onChange={(value) => setRecord({ ...record, correccionesMejoras: value })} />
      <TextArea label="Actividades de operacion" value={record.actividadesOperacion} onChange={(value) => setRecord({ ...record, actividadesOperacion: value })} />
      <TextArea label="Observaciones" value={record.observaciones} onChange={(value) => setRecord({ ...record, observaciones: value })} />
    </>
  )
}

function EditableRows<T extends Record<string, string | number>>({
  title,
  rows,
  onAdd,
  onChange,
  fields,
}: {
  title: string
  rows: T[]
  onAdd: () => void
  onChange: (rows: T[]) => void
  fields: RowField<T>[]
}) {
  function update(index: number, key: keyof T, value: string) {
    const next = rows.map((row, rowIndex) =>
      rowIndex === index ? { ...row, [key]: typeof row[key] === 'number' ? Number(value) : value } : row,
    )
    onChange(next)
  }

  return (
    <section className="entry-section">
      <div className="section-heading">
        <h2>{title}</h2>
        <button type="button" onClick={onAdd}>
          <Plus size={16} /> Renglon
        </button>
      </div>
      {rows.map((row, index) => (
        <div className="row-card" key={index}>
          <div className="row-title">
            <strong>Registro {index + 1}</strong>
            {rows.length > 1 && (
              <button type="button" className="icon-button danger" onClick={() => onChange(rows.filter((_, rowIndex) => rowIndex !== index))}>
                <Trash2 size={16} />
              </button>
            )}
          </div>
          <div className="form-grid dense">
            {fields.map((field) => (
              <Field
                key={String(field.key)}
                label={field.label}
                type={field.type ?? (typeof row[field.key] === 'number' ? 'number' : 'text')}
                value={row[field.key]}
                options={field.options}
                onChange={(value) => update(index, field.key, value)}
              />
            ))}
          </div>
        </div>
      ))}
    </section>
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

function loadApiUrl() {
  return localStorage.getItem(API_URL_KEY) ?? import.meta.env.VITE_API_URL ?? ''
}

function resolveApiBase(apiUrl: string) {
  return apiUrl.trim().replace(/\/$/, '')
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

const drillFields = (options: string[]): RowField<DrillRow>[] => [
  { key: 'equipo', label: 'Equipo', options },
  { key: 'operador', label: 'Operador' },
  { key: 'ayudante', label: 'Ayudante' },
  { key: 'nivelObra', label: 'Nivel / Obra' },
  { key: 'rpaCfte', label: 'Rpa/Cfte' },
  { key: 'longitud', label: 'Longitud' },
  { key: 'cuele', label: 'Cuele' },
  { key: 'desarrollo', label: 'Desb/desc/corte' },
  { key: 'barrenosDados', label: 'Barrenos dados' },
  { key: 'barrenosCargados', label: 'Barrenos cargados' },
  { key: 'metrosDados', label: 'Metros dados' },
  { key: 'horasServicio', label: 'Hrs servicio' },
  { key: 'zanco', label: 'Zanco' },
  { key: 'cople', label: 'Cople' },
  { key: 'barra', label: 'Barra' },
  { key: 'broca', label: 'Broca' },
  { key: 'horometroDieselInicial', label: 'Diesel inicial' },
  { key: 'horometroDieselFinal', label: 'Diesel final' },
  { key: 'horometroElectInicial', label: 'Elect. inicial' },
  { key: 'horometroElectFinal', label: 'Elect. final' },
]

const blastFields: RowField<BlastRow>[] = [
  { key: 'obra', label: 'Obra' },
  { key: 'oficial', label: 'Oficial voladuras' },
  { key: 'ayudante', label: 'Ayudante' },
  { key: 'rpaCfte', label: 'Rpa/Cfte' },
  { key: 'longitud', label: 'Longitud' },
  { key: 'cuele', label: 'Cuele' },
  { key: 'desarrollo', label: 'Desb/desc/corte' },
  { key: 'barrenosPegados', label: 'Barrenos pegados' },
  { key: 'metrosPegados', label: 'Metros pegados' },
  { key: 'horasServicio', label: 'Hrs servicio' },
  { key: 'anfoInicial', label: 'ANFO inicial' },
  { key: 'anfoFinal', label: 'ANFO final' },
  { key: 'observaciones', label: 'Observaciones' },
]

const haulFields: RowField<HaulRow>[] = [
  { key: 'equipo', label: 'Equipo', options: scoopEquipoOptions },
  { key: 'operador', label: 'Operador' },
  { key: 'nivelObra', label: 'Nivel / Obra' },
  { key: 'destino', label: 'Destino' },
  { key: 'rezagado', label: 'Rezagado' },
  { key: 'traspaleo', label: 'Traspaleo' },
  { key: 'limpia', label: 'Limpia' },
  { key: 'balastreo', label: 'Balastreo' },
  { key: 'planilla', label: 'Planilla' },
  { key: 'relleno', label: 'Relleno' },
  { key: 'camiones', label: 'Camiones' },
  { key: 'horometroInicial', label: 'Hor. inicial' },
  { key: 'horometroFinal', label: 'Hor. final' },
  { key: 'diesel', label: 'Diesel' },
  { key: 'observaciones', label: 'Observaciones' },
]

const retroFields: RowField<RetroRow>[] = [
  { key: 'equipo', label: 'Equipo', options: retroEquipoOptions },
  { key: 'operador', label: 'Operador' },
  { key: 'nivelObra', label: 'Nivel / Obra' },
  { key: 'amacice', label: 'Amacice' },
  { key: 'reAmacice', label: 'Re-amacice' },
  { key: 'tableo', label: 'Tableo' },
  { key: 'limpia', label: 'Limpia' },
  { key: 'balastreo', label: 'Balastreo' },
  { key: 'mtAcequia', label: 'MT Acequia' },
  { key: 'horometroInicial', label: 'Hor. inicial' },
  { key: 'horometroFinal', label: 'Hor. final' },
  { key: 'diesel', label: 'Diesel' },
  { key: 'observaciones', label: 'Observaciones' },
]

const activityLabels = {
  polvorero: 'Polvorero',
  choferCamion: 'Chofer camion personal',
  choferPipa: 'Chofer pipa',
  bobCat: 'Bob cat',
  bombeo: 'Bombeo',
  servicios: 'Servicios',
}

export default App
