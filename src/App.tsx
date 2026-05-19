import { useEffect, useMemo, useRef, useState } from 'react'
import type { ChangeEvent, FormEvent, ReactNode } from 'react'
import {
  Accessibility,
  Activity,
  BarChart3,
  Bell,
  ClipboardCheck,
  Drill,
  Edit3,
  FileSpreadsheet,
  FileDown,
  HardHat,
  History,
  Home,
  LayoutDashboard,
  Menu,
  MessageSquare,
  MoreVertical,
  Mountain,
  Pickaxe,
  Plus,
  RefreshCw,
  Save,
  Send,
  ShieldCheck,
  Trash2,
  Truck,
  Upload,
  X,
} from 'lucide-react'
import { Capacitor } from '@capacitor/core'
import ExcelJS from 'exceljs'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'
import type { PDFPage as PdfPage, PDFFont as PdfFont } from 'pdf-lib'
import {
  jumboEquipoOptions,
  retroEquipoOptions,
  scoopEquipoOptions,
} from './data/equipos'
import './App.css'

type Shift = '1' | '2'
type AppSection = 'home' | 'captura' | 'historial' | 'dashboard'
type CaptureMode = 'modulo' | 'turno'
type CaptureLayout = 'asistente' | 'formulario'
type RecordType = 'barrenacion' | 'rezagado' | 'seguridad'
type BarrenacionActivity = 'jumbo' | 'maquinaPierna' | 'voladura'
type RezagadoEquipment = 'scoop' | 'retro'
type BarrenacionAssistantStep = 'actividad' | 'equipo' | 'produccion' | 'detalles'
type RezagadoAssistantStep = 'equipo' | 'trabajo' | 'produccion' | 'cierre'
type HaulActivity = 'rezagado' | 'traspaleo' | 'limpia' | 'balastreo' | 'planilla' | 'relleno'
type RetroActivity = 'amacice' | 'reAmacice' | 'tableo' | 'limpia' | 'balastreo' | 'mtAcequia'
type ChatMessageType = 'aviso' | 'mensaje' | 'urgente'

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
  selectedActivities?: HaulActivity[]
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
  selectedActivities?: RetroActivity[]
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
  activeActivity?: BarrenacionActivity
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
  activeEquipment?: RezagadoEquipment
  activeHaulActivity?: HaulActivity
  activeRetroActivity?: RetroActivity
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
type ShiftBase = Pick<BaseRecord, 'supervisor' | 'fecha' | 'turno' | 'unidad'>
type TurnoModules = Record<RecordType, boolean>
type CaptureDraft = {
  captureMode: CaptureMode
  captureLayout: CaptureLayout
  formType: RecordType
  turnoBase: ShiftBase
  turnoModules: TurnoModules
  barrenacion: BarrenacionRecord
  rezagado: RezagadoRecord
  seguridad: SeguridadRecord
  editingId: string | null
  savedAt: string
}
type CaptureReview = {
  title: string
  records: MineRecord[]
  warnings: string[]
  errors: string[]
  kpis: ReturnType<typeof computeKpis>
}
type EquipmentFavorites = {
  jumbo: string[]
  scoop: string[]
  retro: string[]
}

type ChatMessage = {
  id: string
  type: ChatMessageType
  author: string
  text: string
  createdAt: string
  updatedAt: string
  syncedAt?: string
}

const STORAGE_KEY = 'mga-bitacora-operaciones-v1'
const CAPTURE_DRAFT_KEY = 'mga-bitacora-capture-draft-v1'
const MESSAGE_STORAGE_KEY = 'mga-bitacora-messages-v1'
const API_URL_KEY = 'mga-bitacora-api-url'
const WHATSAPP_NUMBER_KEY = 'mga-bitacora-whatsapp-number'
const OPERATOR_MODE_KEY = 'mga-bitacora-operator-mode'
const DEFAULT_API_URL = 'https://mga-bitacora-mina.onrender.com'
const BARRENACION_TEMPLATE = '/templates/barrenacion-voladuras.pdf'
const REZAGADO_TEMPLATE = '/templates/rezagado.pdf'
const barrenacionActivities = ['jumbo', 'maquinaPierna', 'voladura'] as const
const rezagadoEquipments = ['scoop', 'retro'] as const
const today = new Date().toISOString().slice(0, 10)

const baseDefaults = {
  supervisor: '',
  fecha: today,
  turno: '1' as Shift,
  unidad: 'Unidad Providencia',
}

const defaultTurnoModules: TurnoModules = {
  barrenacion: true,
  rezagado: true,
  seguridad: false,
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
  selectedActivities: ['rezagado'],
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
  selectedActivities: ['amacice'],
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
    activeActivity: 'jumbo',
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
    activeEquipment: 'scoop',
    activeHaulActivity: 'rezagado',
    activeRetroActivity: 'amacice',
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
  const initialDraft = useMemo(loadCaptureDraft, [])
  const [section, setSection] = useState<AppSection>(getInitialSection)
  const [captureMode, setCaptureMode] = useState<CaptureMode>(initialDraft?.captureMode ?? 'modulo')
  const [captureLayout, setCaptureLayout] = useState<CaptureLayout>(initialDraft?.captureLayout ?? 'asistente')
  const [formType, setFormType] = useState<RecordType>(initialDraft?.formType ?? 'barrenacion')
  const [turnoBase, setTurnoBase] = useState<ShiftBase>(initialDraft?.turnoBase ?? { ...baseDefaults })
  const [turnoModules, setTurnoModules] = useState<TurnoModules>(initialDraft?.turnoModules ?? { ...defaultTurnoModules })
  const [barrenacion, setBarrenacion] = useState<BarrenacionRecord>(() => initialDraft?.barrenacion ?? makeBarrenacion())
  const [rezagado, setRezagado] = useState<RezagadoRecord>(() => initialDraft?.rezagado ?? makeRezagado())
  const [seguridad, setSeguridad] = useState<SeguridadRecord>(() => initialDraft?.seguridad ?? makeSeguridad())
  const [records, setRecords] = useState<MineRecord[]>(loadRecords)
  const [chatMessages, setChatMessages] = useState<ChatMessage[]>(loadChatMessages)
  const [chatType, setChatType] = useState<ChatMessageType>('aviso')
  const [chatAuthor, setChatAuthor] = useState('')
  const [chatText, setChatText] = useState('')
  const [whatsNumber, setWhatsNumber] = useState(loadWhatsAppNumber)
  const [query, setQuery] = useState('')
  const [message, setMessage] = useState(() => initialDraft ? `Borrador recuperado: ${formatDateTime(initialDraft.savedAt)}.` : '')
  const [editingId, setEditingId] = useState<string | null>(initialDraft?.editingId ?? null)
  const [syncing, setSyncing] = useState(false)
  const [reviewOpen, setReviewOpen] = useState(false)
  const [apiUrl, setApiUrl] = useState(loadApiUrl)
  const [operatorMode, setOperatorMode] = useState(loadOperatorMode)
  const [drawerOpen, setDrawerOpen] = useState(false)
  const [actionsOpen, setActionsOpen] = useState(false)
  const [chatPanelOpen, setChatPanelOpen] = useState(false)
  const reportRef = useRef<HTMLDivElement>(null)
  const isApk = Capacitor.isNativePlatform()

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
  const pendingMessages = useMemo(
    () => chatMessages.filter((item) => item.updatedAt !== item.syncedAt).length,
    [chatMessages],
  )
  const pendingTotal = pendingSync + pendingMessages
  const latestMessages = useMemo(() => chatMessages.slice(0, 8), [chatMessages])
  const equipmentFavorites = useMemo(() => buildEquipmentFavorites(visibleRecords), [visibleRecords])
  const captureReview = useMemo(
    () => buildCaptureReview(captureMode, formType, turnoBase, turnoModules, barrenacion, rezagado, seguridad, editingId),
    [captureMode, formType, turnoBase, turnoModules, barrenacion, rezagado, seguridad, editingId],
  )

  useEffect(() => {
    localStorage.setItem(API_URL_KEY, apiUrl)
  }, [apiUrl])

  useEffect(() => {
    localStorage.setItem(WHATSAPP_NUMBER_KEY, whatsNumber)
  }, [whatsNumber])

  useEffect(() => {
    localStorage.setItem(OPERATOR_MODE_KEY, operatorMode ? '1' : '0')
  }, [operatorMode])

  useEffect(() => {
    const draft: CaptureDraft = {
      captureMode,
      captureLayout,
      formType,
      turnoBase,
      turnoModules,
      barrenacion,
      rezagado,
      seguridad,
      editingId,
      savedAt: nowIso(),
    }
    if (hasCaptureDraftContent(draft)) {
      localStorage.setItem(CAPTURE_DRAFT_KEY, JSON.stringify(draft))
      return
    }
    localStorage.removeItem(CAPTURE_DRAFT_KEY)
  }, [captureMode, captureLayout, formType, turnoBase, turnoModules, barrenacion, rezagado, seguridad, editingId])

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setDrawerOpen(false)
        setActionsOpen(false)
        setChatPanelOpen(false)
        setReviewOpen(false)
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [])

  function persist(next: MineRecord[]) {
    const normalized = mergeRecords(next.map(normalizeRecord))
    setRecords(normalized)
    localStorage.setItem(STORAGE_KEY, JSON.stringify(normalized))
    return normalized
  }

  function persistMessages(next: ChatMessage[]) {
    const normalized = mergeChatMessages(next.map(normalizeChatMessage))
    setChatMessages(normalized)
    localStorage.setItem(MESSAGE_STORAGE_KEY, JSON.stringify(normalized))
    return normalized
  }

  function publishChatMessage(options: { openWhatsApp?: boolean } = {}) {
    const text = chatText.trim()
    if (!text) return
    const whatsappTarget = normalizeWhatsAppNumber(whatsNumber)
    if (options.openWhatsApp && !whatsappTarget) {
      setMessage('Agrega un numero de WhatsApp con lada para enviar el aviso.')
      return
    }
    const now = nowIso()
    const nextMessage = normalizeChatMessage({
      id: crypto.randomUUID(),
      type: chatType,
      author: chatAuthor.trim() || 'Operacion',
      text,
      createdAt: now,
      updatedAt: now,
    })
    persistMessages([nextMessage, ...chatMessages])
    setChatText('')
    setMessage(options.openWhatsApp ? 'Aviso guardado localmente. Abriendo WhatsApp.' : 'Mensaje guardado localmente. Usa Sincronizar para enviarlo a red.')
    if (options.openWhatsApp && whatsappTarget) openWhatsAppMessage(nextMessage, whatsappTarget)
  }

  function sendChatMessage(event: FormEvent) {
    event.preventDefault()
    publishChatMessage()
  }

  function sendWhatsAppNotice() {
    publishChatMessage({ openWhatsApp: true })
  }

  function openWhatsAppMessage(item: ChatMessage, phone: string) {
    const publicUrl = resolveApiBase(apiUrl)
    const text = [
      'MGA Bitacora Mina',
      `${labelChatType(item.type)}: ${item.text}`,
      `Envia: ${item.author}`,
      `Fecha: ${formatDateTime(item.createdAt)}`,
      `Panel: ${publicUrl}`,
    ].join('\n')
    window.open(`https://wa.me/${phone}?text=${encodeURIComponent(text)}`, '_blank', 'noopener,noreferrer')
  }

  async function syncMessages(options: { silent?: boolean; sourceMessages?: ChatMessage[] } = {}) {
    const apiBase = resolveApiBase(apiUrl)
    if (!apiBase || (typeof navigator !== 'undefined' && !navigator.onLine)) {
      if (!options.silent) setMessage('Sin internet. Los mensajes quedan guardados localmente; vuelve a presionar Sincronizar cuando haya red.')
      return
    }
    try {
      const localMessages = (options.sourceMessages ?? chatMessages).map(normalizeChatMessage)
      const response = await fetch(`${apiBase}/api/messages/sync`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ messages: localMessages }),
      })
      if (!response.ok) throw new Error(`HTTP ${response.status}`)
      const payload = (await response.json()) as { messages: ChatMessage[] }
      persistMessages(
        mergeChatMessages([...localMessages, ...payload.messages.map(normalizeChatMessage)]).map((item) => ({
          ...item,
          syncedAt: item.updatedAt,
        })),
      )
      if (!options.silent) setMessage('Centro de mensajes sincronizado.')
    } catch {
      if (!options.silent) setMessage('No se pudieron sincronizar los mensajes. Siguen guardados localmente.')
    }
  }

  async function syncAll() {
    await syncRecords()
    await syncMessages()
  }

  function saveRecord(event: FormEvent) {
    event.preventDefault()
    setReviewOpen(true)
  }

  function commitReviewedSave() {
    if (captureReview.errors.length > 0) return
    setReviewOpen(false)
    if (captureMode === 'turno' && !editingId) {
      saveTurnoCompleto()
      return
    }
    saveModuloRecord()
  }

  function saveModuloRecord() {
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
    setMessage(editingId ? 'Captura actualizada en el historial local.' : 'Registro guardado en el historial local de este dispositivo.')
  }

  function saveTurnoCompleto() {
    const selectedRecords: MineRecord[] = []
    const updatedAt = nowIso()
    if (turnoModules.barrenacion) selectedRecords.push(applyShiftBase(barrenacion, turnoBase, updatedAt))
    if (turnoModules.rezagado) selectedRecords.push(applyShiftBase(rezagado, turnoBase, updatedAt))
    if (turnoModules.seguridad) selectedRecords.push(applyShiftBase(seguridad, turnoBase, updatedAt))
    if (selectedRecords.length === 0) {
      setMessage('Selecciona al menos un modulo del turno antes de guardar.')
      return
    }
    persist([...selectedRecords, ...records])
    resetTurnoCompleto()
    setCaptureMode('modulo')
    setSection('historial')
    setMessage(`Turno completo guardado en historial local con ${selectedRecords.length} registro(s).`)
  }

  function resetForm(type: RecordType) {
    if (type === 'barrenacion') setBarrenacion(makeBarrenacion())
    if (type === 'rezagado') setRezagado(makeRezagado())
    if (type === 'seguridad') setSeguridad(makeSeguridad())
  }

  function resetTurnoCompleto() {
    setTurnoBase({ ...baseDefaults, fecha: new Date().toISOString().slice(0, 10) })
    setTurnoModules({ ...defaultTurnoModules })
    setBarrenacion(makeBarrenacion())
    setRezagado(makeRezagado())
    setSeguridad(makeSeguridad())
  }

  function startEdit(record: MineRecord) {
    const cleanRecord = { ...record, deletedAt: undefined } as MineRecord
    setCaptureMode('modulo')
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
    setCaptureMode('modulo')
    setFormType(type)
  }

  function closeMenus() {
    setDrawerOpen(false)
    setActionsOpen(false)
  }

  function goHome() {
    setSection('home')
    closeMenus()
  }

  function goToCapture() {
    setCaptureMode('modulo')
    setSection('captura')
    closeMenus()
  }

  function goToHistory() {
    setSection('historial')
    closeMenus()
  }

  function openCapture(type: RecordType) {
    selectFormType(type)
    setCaptureMode('modulo')
    setSection('captura')
    closeMenus()
  }

  function startTurnoCompleto() {
    if (editingId) {
      setMessage('Guarda o cancela la edicion antes de abrir turno completo.')
      return
    }
    if (captureMode !== 'turno' || section !== 'captura') resetTurnoCompleto()
    setCaptureMode('turno')
    setCaptureLayout('asistente')
    setSection('captura')
    closeMenus()
    setMessage('Turno completo listo. Captura los datos generales una sola vez.')
  }

  function goToDashboard() {
    if (isApk) {
      setSection('historial')
      closeMenus()
      return
    }
    setSection('dashboard')
    closeMenus()
  }

  function startNewCapture() {
    if (captureMode === 'turno') resetTurnoCompleto()
    else resetForm(formType)
    setEditingId(null)
    setSection('captura')
    closeMenus()
    setMessage('Nueva captura lista.')
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
    setMessage('Captura eliminada del historial local. Usa Sincronizar para reflejarlo en red.')
  }

  async function importJson(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return
    const imported = JSON.parse(await file.text()) as MineRecord[]
    persist([...imported.map(normalizeRecord), ...records])
    setMessage('Datos importados al historial local.')
    closeMenus()
    event.target.value = ''
  }

  async function syncRecords(options: { silent?: boolean; sourceRecords?: MineRecord[] } = {}) {
    const apiBase = resolveApiBase(apiUrl)
    if (!apiBase || (typeof navigator !== 'undefined' && !navigator.onLine)) {
      if (!options.silent) setMessage('Sin internet. La bitacora queda guardada en este dispositivo; vuelve a presionar Sincronizar cuando haya red.')
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
      if (!options.silent) setMessage('No se pudo sincronizar. Los registros siguen guardados localmente.')
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

  async function exportRecordPdf(record: MineRecord) {
    if (record.type === 'seguridad') {
      setMessage('El formato PDF adjunto aplica a barrenacion/voladuras y rezagado.')
      return
    }
    const bytes = record.type === 'barrenacion'
      ? await buildBarrenacionPdf(record)
      : await buildRezagadoPdf(record)
    downloadBlob(
      bytes,
      `${record.type}-${record.fecha}-turno-${record.turno}.pdf`,
      'application/pdf',
    )
  }

  async function exportRecordExcel(record: MineRecord) {
    if (record.type === 'seguridad') {
      setMessage('El formato Excel adjunto aplica a barrenacion/voladuras y rezagado.')
      return
    }
    const bytes = record.type === 'barrenacion'
      ? await buildBarrenacionExcel(record)
      : await buildRezagadoExcel(record)
    downloadBlob(
      bytes,
      `${record.type}-${record.fecha}-turno-${record.turno}.xlsx`,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
  }

  function toggleOperatorMode() {
    setOperatorMode((current) => !current)
    setMessage(operatorMode ? 'Modo operador desactivado.' : 'Modo operador activado: controles mas grandes para captura en campo.')
  }

  return (
    <main className={`app-shell ${operatorMode ? 'operator-mode' : ''} ${section === 'captura' ? 'capture-mode' : section === 'dashboard' || section === 'historial' ? 'review-mode' : 'home-mode'}`}>
      <header className="topbar">
        <button
          className="chrome-icon"
          aria-expanded={drawerOpen}
          aria-label="Abrir menu"
          type="button"
          onClick={() => {
            setDrawerOpen((open) => !open)
            setActionsOpen(false)
          }}
        >
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
          <button className={section === 'home' ? 'active' : ''} onClick={goHome}>
            <Home size={18} /> Inicio
          </button>
          <button className={section === 'captura' ? 'active' : ''} onClick={goToCapture}>
            <ClipboardCheck size={18} /> Captura
          </button>
          <button className={section === 'historial' ? 'active' : ''} onClick={goToHistory}>
            <History size={18} /> Historial
          </button>
          {!isApk && (
            <button
              className={section === 'dashboard' ? 'active' : ''}
              onClick={goToDashboard}
            >
              <LayoutDashboard size={18} /> Revision web
            </button>
          )}
        </nav>
        <button className="sync-button" onClick={() => void syncAll()} disabled={syncing}>
          <RefreshCw size={18} className={syncing ? 'spin' : ''} />
          {syncing ? 'Conectando' : pendingTotal > 0 ? 'Pendiente' : 'Al dia'}
          {pendingTotal > 0 && <span>{pendingTotal}</span>}
        </button>
        <button
          className={`operator-toggle ${operatorMode ? 'active' : ''}`}
          type="button"
          aria-pressed={operatorMode}
          onClick={toggleOperatorMode}
        >
          <Accessibility size={18} />
          <span>Operador</span>
        </button>
        <button
          className="chrome-icon"
          aria-expanded={actionsOpen}
          aria-label="Abrir acciones"
          type="button"
          onClick={() => {
            setActionsOpen((open) => !open)
            setDrawerOpen(false)
          }}
        >
          <MoreVertical size={21} />
        </button>
      </header>

      {(drawerOpen || actionsOpen) && <button className="menu-backdrop" type="button" aria-label="Cerrar menu" onClick={closeMenus} />}

      {drawerOpen && (
        <aside className="app-drawer open">
          <div className="drawer-brand">
            <img src="/mga-logo.jfif" alt="MGA" />
            <div>
              <span>MGA Operaciones Mina</span>
              <strong>Panel de captura</strong>
            </div>
          </div>
          <button className={section === 'home' ? 'active' : ''} type="button" onClick={goHome}>
            <Home size={20} />
            <span>Inicio</span>
            <small>Menu de modulos</small>
          </button>
          <button className={section === 'captura' ? 'active' : ''} type="button" onClick={goToCapture}>
            <ClipboardCheck size={20} />
            <span>Captura de campo</span>
            <small>Registro rapido offline</small>
          </button>
          <button className={captureMode === 'turno' && section === 'captura' ? 'active' : ''} type="button" onClick={startTurnoCompleto}>
            <Mountain size={20} />
            <span>Turno completo</span>
            <small>Una bitacora con varios modulos</small>
          </button>
          <button className={section === 'historial' ? 'active' : ''} type="button" onClick={goToHistory}>
            <History size={20} />
            <span>Historial local</span>
            <small>Bitacora guardada en el equipo</small>
          </button>
          {!isApk && (
            <button className={section === 'dashboard' ? 'active' : ''} type="button" onClick={goToDashboard}>
              <LayoutDashboard size={20} />
              <span>Revision web</span>
              <small>KPI, reportes y descargas</small>
            </button>
          )}
          <button type="button" onClick={startNewCapture}>
            <Edit3 size={20} />
            <span>Nueva captura</span>
            <small>Limpia el formulario actual</small>
          </button>
          <button className={operatorMode ? 'active' : ''} type="button" onClick={toggleOperatorMode}>
            <Accessibility size={20} />
            <span>Modo operador</span>
            <small>{operatorMode ? 'Controles grandes activos' : 'Botones y texto mas grandes'}</small>
          </button>
          <button type="button" onClick={() => {
            setChatPanelOpen(true)
            closeMenus()
          }}>
            <MessageSquare size={20} />
            <span>Centro de mensajes</span>
            <small>{pendingMessages > 0 ? `${pendingMessages} pendientes` : 'Avisos y comunicados'}</small>
          </button>
          <button type="button" onClick={() => {
            closeMenus()
            void syncAll()
          }}>
            <RefreshCw size={20} className={syncing ? 'spin' : ''} />
            <span>Sincronizar</span>
            <small>{pendingTotal > 0 ? `${pendingTotal} pendientes` : 'Datos al dia'}</small>
          </button>
        </aside>
      )}

      {actionsOpen && (
        <div className="actions-menu open">
          <button type="button" onClick={goHome}>
            <Home size={18} /> Inicio
          </button>
          <button type="button" onClick={startNewCapture}>
            <Edit3 size={18} /> Nueva captura
          </button>
          <button type="button" onClick={startTurnoCompleto}>
            <Mountain size={18} /> Turno completo
          </button>
          <button type="button" onClick={toggleOperatorMode}>
            <Accessibility size={18} /> {operatorMode ? 'Modo normal' : 'Modo operador'}
          </button>
          <button type="button" onClick={goToHistory}>
            <History size={18} /> Historial local
          </button>
          <button type="button" onClick={() => {
            closeMenus()
            void syncAll()
          }}>
            <RefreshCw size={18} className={syncing ? 'spin' : ''} /> Sincronizar
          </button>
          {!isApk && (
            <button type="button" onClick={goToDashboard}>
              <LayoutDashboard size={18} /> Revision web
            </button>
          )}
          <button type="button" onClick={() => {
            setChatPanelOpen(true)
            closeMenus()
          }}>
            <MessageSquare size={18} /> Centro de mensajes
          </button>
          <label className="file-button menu-file">
            <Upload size={18} /> Importar respaldo
            <input type="file" accept="application/json" onChange={importJson} />
          </label>
          {section === 'dashboard' && (
            <button type="button" onClick={() => {
              closeMenus()
              void exportPdf()
            }}>
              <FileDown size={18} /> PDF KPI
            </button>
          )}
          {editingId && (
            <button type="button" onClick={() => {
              closeMenus()
              cancelEdit()
            }}>
              <X size={18} /> Cancelar edicion
            </button>
          )}
        </div>
      )}

      {message && <div className="toast" role="status" aria-live="polite">{message}</div>}
      {editingId && section !== 'captura' && (
        <div className="edit-banner">
          <span>Hay una captura en edicion.</span>
          <button type="button" onClick={goToCapture}>
            Continuar
          </button>
          <button type="button" onClick={cancelEdit}>
            Cancelar
          </button>
        </div>
      )}
      {reviewOpen && (
        <SaveReviewModal
          review={captureReview}
          onCancel={() => setReviewOpen(false)}
          onConfirm={commitReviewedSave}
        />
      )}

      {section === 'home' ? (
        <HomeScreen
          pendingMessages={pendingMessages}
          pendingSync={pendingTotal}
          recordCounts={recordCounts}
          showDashboard={!isApk}
          syncing={syncing}
          onDashboard={goToDashboard}
          onHistory={goToHistory}
          onMessages={() => setChatPanelOpen(true)}
          onNewCapture={startNewCapture}
          onOpenCapture={openCapture}
          onSync={() => void syncAll()}
          onTurnoCompleto={startTurnoCompleto}
        />
      ) : section === 'captura' ? (
        <form id="capture-form" className="workspace" onSubmit={saveRecord}>
          <aside className="side-panel">
            <div className="capture-badge">
              <Mountain size={18} /> Operacion mina
            </div>
            <h1>{captureMode === 'turno' ? 'Turno completo' : editingId ? 'Editar captura' : 'Nueva captura'}</h1>
            <p>
              {captureMode === 'turno'
                ? 'Captura los datos generales una sola vez y agrega los modulos trabajados durante el turno.'
                : isApk
                ? 'Captura un evento de campo con los datos clave. Todo queda en el historial local y se envia a red solo al sincronizar.'
                : 'Captura un evento de campo con los datos clave. El reporte completo se arma en la revision web.'}
            </p>
            <div className="connection-card">
              <span>Sincronizacion manual</span>
              <strong>{syncing ? 'Sincronizando' : pendingSync > 0 ? `${pendingSync} por enviar` : 'Historial local al dia'}</strong>
            </div>
            <div className="record-tabs">
              <button
                type="button"
                className={`module-card module-slate ${captureMode === 'turno' ? 'active' : ''}`}
                onClick={startTurnoCompleto}
              >
                <span className="module-icon"><Mountain size={24} /></span>
                <span className="module-title">Turno</span>
                <span className="module-subtitle">Completo</span>
                <span className="module-ring"><strong>{recordCounts.total}</strong><small>hist.</small></span>
              </button>
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
              {!isApk && (
                <button
                  type="button"
                  className="module-card module-green"
                  onClick={() => {
                    setSection('dashboard')
                  }}
                >
                  <span className="module-icon"><Pickaxe size={24} /></span>
                  <span className="module-title">Revision</span>
                  <span className="module-subtitle">KPI / reportes</span>
                  <span className="module-ring"><strong>{recordCounts.total}</strong><small>total</small></span>
                </button>
              )}
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
                      <strong>{record.fecha} - Turno {record.turno}</strong>
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
            {captureMode === 'turno' && !editingId ? (
              <TurnoCompletoForm
                barrenacion={barrenacion}
                base={turnoBase}
                captureLayout={captureLayout}
                equipmentFavorites={equipmentFavorites}
                modules={turnoModules}
                rezagado={rezagado}
                seguridad={seguridad}
                setBarrenacion={setBarrenacion}
                setBase={setTurnoBase}
                setCaptureLayout={setCaptureLayout}
                setModules={setTurnoModules}
                setRezagado={setRezagado}
                setSeguridad={setSeguridad}
              />
            ) : (
              <>
                <CaptureFlowNav formType={formType} />
                {formType === 'barrenacion' && (
                  <BarrenacionForm record={barrenacion} setRecord={setBarrenacion} />
                )}
                {formType === 'rezagado' && <RezagadoForm record={rezagado} setRecord={setRezagado} />}
                {formType === 'seguridad' && <SeguridadForm record={seguridad} setRecord={setSeguridad} />}
              </>
            )}
          </section>
        </form>
      ) : section === 'historial' ? (
        <HistoryScreen
          filtered={filtered}
          kpis={kpis}
          pendingSync={pendingSync}
          query={query}
          syncing={syncing}
          totalRecords={recordCounts.total}
          onDelete={removeRecord}
          onEdit={startEdit}
          onExportExcel={exportRecordExcel}
          onExportPdf={exportRecordPdf}
          onNewCapture={startNewCapture}
          onQueryChange={setQuery}
          onSync={() => void syncAll()}
        />
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
              <button onClick={() => void syncAll()} disabled={syncing}>
                <RefreshCw size={18} className={syncing ? 'spin' : ''} /> Sincronizar
              </button>
              <button onClick={exportPdf}>
                <FileDown size={18} /> PDF KPI
              </button>
            </div>
          </div>

          <MessageCenter
            chatAuthor={chatAuthor}
            chatText={chatText}
            chatType={chatType}
            latestMessages={latestMessages}
            pendingMessages={pendingMessages}
            setChatAuthor={setChatAuthor}
            setChatText={setChatText}
            setChatType={setChatType}
            setWhatsNumber={setWhatsNumber}
            syncing={syncing}
            whatsNumber={whatsNumber}
            onSubmit={sendChatMessage}
            onSync={() => void syncMessages()}
            onWhatsApp={sendWhatsAppNotice}
          />

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
                        {record.type !== 'seguridad' && (
                          <>
                            <button className="icon-button" onClick={() => void exportRecordPdf(record)} aria-label="Descargar PDF formato">
                              <FileDown size={16} />
                            </button>
                            <button className="icon-button" onClick={() => void exportRecordExcel(record)} aria-label="Descargar Excel formato">
                              <FileSpreadsheet size={16} />
                            </button>
                          </>
                        )}
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
      {chatPanelOpen && (
        <>
          <button className="chat-panel-backdrop" type="button" aria-label="Cerrar centro de mensajes" onClick={() => setChatPanelOpen(false)} />
          <aside className="floating-message-panel" aria-label="Centro de mensajes flotante">
            <button className="floating-panel-close" type="button" onClick={() => setChatPanelOpen(false)} aria-label="Cerrar centro de mensajes">
              <X size={18} />
            </button>
            <MessageCenter
              chatAuthor={chatAuthor}
              chatText={chatText}
              chatType={chatType}
              latestMessages={latestMessages}
              pendingMessages={pendingMessages}
              setChatAuthor={setChatAuthor}
              setChatText={setChatText}
              setChatType={setChatType}
              setWhatsNumber={setWhatsNumber}
              syncing={syncing}
              whatsNumber={whatsNumber}
              onSubmit={sendChatMessage}
              onSync={() => void syncMessages()}
              onWhatsApp={sendWhatsAppNotice}
            />
          </aside>
        </>
      )}
      <button
        className="floating-chat-button"
        type="button"
        aria-label="Abrir centro de mensajes"
        onClick={() => setChatPanelOpen(true)}
      >
        <MessageSquare size={24} />
        {pendingMessages > 0 && <span>{pendingMessages}</span>}
      </button>
      <nav className="mobile-bottom-nav" aria-label="Acciones rapidas">
        <button className={section === 'home' ? 'active' : ''} type="button" onClick={goHome}>
          <Home size={20} />
          <span>Inicio</span>
        </button>
        {section === 'captura' ? (
          <button type="submit" form="capture-form">
            <Save size={20} />
            <span>Guardar</span>
          </button>
        ) : (
          <button type="button" onClick={goToCapture}>
            <ClipboardCheck size={20} />
            <span>Captura</span>
          </button>
        )}
        <button type="button" onClick={() => setChatPanelOpen(true)}>
          <MessageSquare size={20} />
          <span>Mensajes</span>
        </button>
        <button className={section === 'historial' ? 'active' : ''} type="button" onClick={goToHistory}>
          <History size={20} />
          <span>Historial</span>
        </button>
        {!isApk && (
          <button className={section === 'dashboard' ? 'active' : ''} type="button" onClick={goToDashboard}>
            <LayoutDashboard size={20} />
            <span>Revision</span>
          </button>
        )}
      </nav>
    </main>
  )
}

function SaveReviewModal({
  review,
  onCancel,
  onConfirm,
}: {
  review: CaptureReview
  onCancel: () => void
  onConfirm: () => void
}) {
  const canSave = review.errors.length === 0
  const reviewStatus = review.errors.length > 0
    ? `${review.errors.length} error(es) por corregir`
    : review.warnings.length > 0
    ? `${review.warnings.length} aviso(s) para revisar`
    : 'Listo para guardar'
  return (
    <>
      <button className="review-modal-backdrop" type="button" aria-label="Cerrar revision" onClick={onCancel} />
      <section className="save-review-modal" role="dialog" aria-modal="true" aria-label="Revision antes de guardar">
        <div className="save-review-head">
          <span>Revision antes de guardar</span>
          <h2>{review.title}</h2>
          <p>Confirma los registros que se guardaran en el historial local.</p>
        </div>

        <div className={`review-status ${review.errors.length > 0 ? 'error' : review.warnings.length > 0 ? 'warning' : 'ready'}`}>
          <strong>{reviewStatus}</strong>
          <span>{canSave ? 'Puedes guardar localmente y sincronizar despues.' : 'Corrige los datos marcados antes de guardar.'}</span>
        </div>

        <div className="review-kpi-grid">
          <article>
            <span>Registros</span>
            <strong>{review.records.length}</strong>
          </article>
          <article>
            <span>Metros barrenados</span>
            <strong>{review.kpis.metrosDados}</strong>
          </article>
          <article>
            <span>Metros pegados</span>
            <strong>{review.kpis.metrosPegados}</strong>
          </article>
          <article>
            <span>Cucharones</span>
            <strong>{review.kpis.rezagado}</strong>
          </article>
        </div>

        <div className="review-record-list">
          {review.records.map((record) => (
            <article key={record.id}>
              <span>{labelType(record.type)}</span>
              <strong>{record.fecha} - Turno {record.turno}</strong>
              <small>{record.supervisor || 'Supervisor pendiente'} - {record.unidad}</small>
            </article>
          ))}
        </div>

        {(review.errors.length > 0 || review.warnings.length > 0) && (
          <div className="review-alerts">
            {review.errors.map((item) => (
              <p className="error" key={item}>{item}</p>
            ))}
            {review.warnings.map((item) => (
              <p key={item}>{item}</p>
            ))}
          </div>
        )}

        <div className="save-review-actions">
          <button type="button" onClick={onCancel}>
            Seguir editando
          </button>
          <button className="primary-action" type="button" onClick={onConfirm} disabled={!canSave}>
            <Save size={18} /> Guardar local
          </button>
        </div>
      </section>
    </>
  )
}

function HomeScreen({
  pendingMessages,
  pendingSync,
  recordCounts,
  showDashboard,
  syncing,
  onDashboard,
  onHistory,
  onMessages,
  onNewCapture,
  onOpenCapture,
  onSync,
  onTurnoCompleto,
}: {
  pendingMessages: number
  pendingSync: number
  recordCounts: { barrenacion: number; rezagado: number; seguridad: number; total: number }
  showDashboard: boolean
  syncing: boolean
  onDashboard: () => void
  onHistory: () => void
  onMessages: () => void
  onNewCapture: () => void
  onOpenCapture: (type: RecordType) => void
  onSync: () => void
  onTurnoCompleto: () => void
}) {
  return (
    <section className="home-screen">
      <div className="home-hero">
        <img src="/mga-logo.jfif" alt="MGA" />
        <div>
          <span>Operacion mina</span>
          <h1>Modulos</h1>
          <p>Elige una tarea y captura solo el bloque necesario.</p>
        </div>
      </div>

      <div className="module-launch-grid" aria-label="Modulos principales">
        <button className="launch-card launch-slate" type="button" onClick={onTurnoCompleto}>
          <span><Mountain size={30} /></span>
          <strong>Turno completo</strong>
          <small>Varios modulos</small>
        </button>
        <button className="launch-card launch-red" type="button" onClick={() => onOpenCapture('barrenacion')}>
          <span><Drill size={30} /></span>
          <strong>Barrenos</strong>
          <small>{recordCounts.barrenacion} registros</small>
        </button>
        <button className="launch-card launch-cyan" type="button" onClick={() => onOpenCapture('rezagado')}>
          <span><Truck size={30} /></span>
          <strong>Rezagado</strong>
          <small>{recordCounts.rezagado} registros</small>
        </button>
        <button className="launch-card launch-amber" type="button" onClick={() => onOpenCapture('seguridad')}>
          <span><ShieldCheck size={30} /></span>
          <strong>Seguridad</strong>
          <small>{recordCounts.seguridad} registros</small>
        </button>
        <button className="launch-card launch-green" type="button" onClick={onHistory}>
          <span><History size={30} /></span>
          <strong>Historial</strong>
          <small>{recordCounts.total} locales</small>
        </button>
        {showDashboard && (
          <button className="launch-card launch-teal" type="button" onClick={onDashboard}>
            <span><LayoutDashboard size={30} /></span>
            <strong>Revision</strong>
            <small>KPI y reportes</small>
          </button>
        )}
        <button className="launch-card launch-violet" type="button" onClick={onMessages}>
          <span><MessageSquare size={30} /></span>
          <strong>Mensajes</strong>
          <small>{pendingMessages > 0 ? `${pendingMessages} pendientes` : 'Avisos'}</small>
        </button>
        <button className="launch-card launch-blue" type="button" onClick={onSync}>
          <span><RefreshCw size={30} className={syncing ? 'spin' : ''} /></span>
          <strong>Sincronizar</strong>
          <small>{pendingSync > 0 ? `${pendingSync} por enviar` : 'Al dia'}</small>
        </button>
        <button className="launch-card launch-slate" type="button" onClick={onNewCapture}>
          <span><Edit3 size={30} /></span>
          <strong>Nueva</strong>
          <small>Limpiar captura</small>
        </button>
        {showDashboard && (
          <button className="launch-card launch-green" type="button" onClick={onDashboard}>
            <span><FileDown size={30} /></span>
            <strong>Formatos</strong>
            <small>PDF / Excel</small>
          </button>
        )}
      </div>

      <div className="home-summary">
        <article>
          <span>Total capturas</span>
          <strong>{recordCounts.total}</strong>
        </article>
        <article>
          <span>Sincronizacion</span>
          <strong>{pendingSync > 0 ? `${pendingSync} pendientes` : 'Al dia'}</strong>
        </article>
      </div>
    </section>
  )
}

function HistoryScreen({
  filtered,
  kpis,
  pendingSync,
  query,
  syncing,
  totalRecords,
  onDelete,
  onEdit,
  onExportExcel,
  onExportPdf,
  onNewCapture,
  onQueryChange,
  onSync,
}: {
  filtered: MineRecord[]
  kpis: ReturnType<typeof computeKpis>
  pendingSync: number
  query: string
  syncing: boolean
  totalRecords: number
  onDelete: (id: string) => void
  onEdit: (record: MineRecord) => void
  onExportExcel: (record: MineRecord) => Promise<void>
  onExportPdf: (record: MineRecord) => Promise<void>
  onNewCapture: () => void
  onQueryChange: (value: string) => void
  onSync: () => void
}) {
  return (
    <section className="history-screen">
      <div className="history-head">
        <div>
          <span>Historial offline</span>
          <h1>Historial de bitacora</h1>
          <p>Todos los registros se quedan en este dispositivo para revision. La red solo se actualiza al presionar Sincronizar.</p>
        </div>
        <button type="button" onClick={onSync} disabled={syncing}>
          <RefreshCw size={18} className={syncing ? 'spin' : ''} />
          {pendingSync > 0 ? `Sincronizar ${pendingSync}` : 'Sincronizar'}
        </button>
      </div>

      <div className="history-tools">
        <input
          aria-label="Buscar historial"
          placeholder="Buscar por fecha, supervisor, turno, unidad..."
          value={query}
          onChange={(event) => onQueryChange(event.target.value)}
        />
        <button type="button" onClick={onNewCapture}>
          <Plus size={18} /> Nueva captura
        </button>
      </div>

      <div className="history-kpis">
        <Kpi icon={<History />} label="Registros locales" value={totalRecords} />
        <Kpi icon={<RefreshCw />} label="Pendientes red" value={pendingSync} />
        <Kpi icon={<BarChart3 />} label="Metros barrenados" value={kpis.metrosDados} />
        <Kpi icon={<HardHat />} label="Cucharones" value={kpis.rezagado} />
      </div>

      <div className="history-list">
        {filtered.length === 0 ? (
          <article className="history-empty">
            <strong>Sin registros para mostrar</strong>
            <p>Cuando guardes una captura, aparecera aqui aunque no haya internet.</p>
          </article>
        ) : (
          filtered.map((record) => (
            <article className="history-record-card" key={record.id}>
              <div>
                <span>{labelType(record.type)}</span>
                <strong>{record.fecha} - Turno {record.turno}</strong>
                <small>{record.supervisor || 'Supervisor pendiente'} - {record.unidad}</small>
              </div>
              <span className={record.updatedAt === record.syncedAt ? 'status-pill synced' : 'status-pill'}>
                {record.updatedAt === record.syncedAt ? 'En red' : 'Local'}
              </span>
              <div className="history-actions">
                {record.type !== 'seguridad' && (
                  <>
                    <button type="button" onClick={() => void onExportPdf(record)}>
                      <FileDown size={16} /> PDF
                    </button>
                    <button type="button" onClick={() => void onExportExcel(record)}>
                      <FileSpreadsheet size={16} /> Excel
                    </button>
                  </>
                )}
                <button type="button" onClick={() => onEdit(record)}>
                  <Edit3 size={16} /> Editar
                </button>
                <button className="danger" type="button" onClick={() => onDelete(record.id)}>
                  <Trash2 size={16} /> Eliminar
                </button>
              </div>
            </article>
          ))
        )}
      </div>
    </section>
  )
}

function CaptureFlowNav({ formType }: { formType: RecordType }) {
  const steps = formType === 'barrenacion'
    ? [
        ['capture-turno', 'Turno', Mountain] as const,
        ['capture-actividad', 'Actividad', Drill] as const,
        ['capture-produccion', 'Produccion', Pickaxe] as const,
        ['capture-cierre', 'Cierre', ClipboardCheck] as const,
      ]
    : formType === 'rezagado'
      ? [
          ['capture-turno', 'Turno', Mountain] as const,
          ['capture-equipo', 'Equipo', Truck] as const,
          ['capture-trabajo', 'Trabajo', Activity] as const,
          ['capture-cierre', 'Cierre', ClipboardCheck] as const,
        ]
      : [
          ['capture-turno', 'Turno', Mountain] as const,
          ['capture-resumen', 'Resumen', ShieldCheck] as const,
          ['capture-hallazgo', 'Hallazgo', ClipboardCheck] as const,
        ]

  return (
    <div className="capture-flow-nav" aria-label="Pasos de captura">
      {steps.map(([target, label, Icon]) => (
        <button key={target} type="button" onClick={() => scrollToCaptureSection(target)}>
          <Icon size={17} />
          <span>{label}</span>
        </button>
      ))}
    </div>
  )
}

function scrollToCaptureSection(target: string) {
  document.getElementById(target)?.scrollIntoView({ behavior: 'smooth', block: 'start' })
}

function MessageCenter({
  chatAuthor,
  chatText,
  chatType,
  latestMessages,
  pendingMessages,
  setChatAuthor,
  setChatText,
  setChatType,
  setWhatsNumber,
  syncing,
  whatsNumber,
  onSubmit,
  onSync,
  onWhatsApp,
}: {
  chatAuthor: string
  chatText: string
  chatType: ChatMessageType
  latestMessages: ChatMessage[]
  pendingMessages: number
  setChatAuthor: (value: string) => void
  setChatText: (value: string) => void
  setChatType: (value: ChatMessageType) => void
  setWhatsNumber: (value: string) => void
  syncing: boolean
  whatsNumber: string
  onSubmit: (event: FormEvent) => void
  onSync: () => void
  onWhatsApp: () => void
}) {
  return (
    <section className="message-center" data-html2canvas-ignore="true">
      <div className="message-head">
        <div>
          <span><MessageSquare size={18} /> Centro de mensajes</span>
          <h2>Avisos y comunicados</h2>
        </div>
        <button type="button" onClick={onSync} disabled={syncing}>
          <RefreshCw size={16} className={syncing ? 'spin' : ''} />
          {pendingMessages > 0 ? `${pendingMessages} pendientes` : 'Al dia'}
        </button>
      </div>
      <form className="message-composer" onSubmit={onSubmit}>
        <div className="message-type-tabs" aria-label="Tipo de mensaje">
          {(['aviso', 'mensaje', 'urgente'] as const).map((type) => (
            <button
              key={type}
              type="button"
              className={chatType === type ? 'active' : ''}
              onClick={() => setChatType(type)}
            >
              {type === 'aviso' && <Bell size={16} />}
              {type === 'mensaje' && <MessageSquare size={16} />}
              {type === 'urgente' && <ShieldCheck size={16} />}
              {labelChatType(type)}
            </button>
          ))}
        </div>
        <input
          aria-label="Nombre de quien envia"
          placeholder="Nombre"
          value={chatAuthor}
          onChange={(event) => setChatAuthor(event.target.value)}
        />
        <input
          aria-label="Numero de WhatsApp"
          inputMode="tel"
          placeholder="WhatsApp con lada"
          value={whatsNumber}
          onChange={(event) => setWhatsNumber(event.target.value)}
        />
        <textarea
          aria-label="Mensaje"
          placeholder="Escribe un aviso para el equipo..."
          value={chatText}
          onChange={(event) => setChatText(event.target.value)}
          rows={3}
        />
        <div className="message-send-actions">
          <button className="send-message" type="submit">
            <Send size={17} /> Publicar
          </button>
          <button className="whatsapp-message" type="button" onClick={onWhatsApp}>
            <MessageSquare size={17} /> WhatsApp
          </button>
        </div>
      </form>
      <div className="message-list">
        {latestMessages.length === 0 ? (
          <p>No hay mensajes todavia.</p>
        ) : (
          latestMessages.map((item) => (
            <article className={`message-item ${item.type}`} key={item.id}>
              <div>
                <span>{labelChatType(item.type)}</span>
                <strong>{item.author}</strong>
                <time>{formatDateTime(item.createdAt)}</time>
              </div>
              <p>{item.text}</p>
            </article>
          ))
        )}
      </div>
    </section>
  )
}

function TurnoCompletoForm({
  barrenacion,
  base,
  captureLayout,
  equipmentFavorites,
  modules,
  rezagado,
  seguridad,
  setBarrenacion,
  setBase,
  setCaptureLayout,
  setModules,
  setRezagado,
  setSeguridad,
}: {
  barrenacion: BarrenacionRecord
  base: ShiftBase
  captureLayout: CaptureLayout
  equipmentFavorites: EquipmentFavorites
  modules: TurnoModules
  rezagado: RezagadoRecord
  seguridad: SeguridadRecord
  setBarrenacion: (record: BarrenacionRecord) => void
  setBase: (base: ShiftBase) => void
  setCaptureLayout: (layout: CaptureLayout) => void
  setModules: (modules: TurnoModules) => void
  setRezagado: (record: RezagadoRecord) => void
  setSeguridad: (record: SeguridadRecord) => void
}) {
  const selectedCount = Object.values(modules).filter(Boolean).length
  const enabledModules = (['barrenacion', 'rezagado', 'seguridad'] as const).filter((type) => modules[type])
  const [assistantModule, setAssistantModule] = useState<RecordType>('barrenacion')
  const activeAssistantModule = modules[assistantModule] ? assistantModule : enabledModules[0] ?? 'barrenacion'

  function toggleModule(type: RecordType) {
    const next = { ...modules, [type]: !modules[type] }
    if (!Object.values(next).some(Boolean)) return
    setModules(next)
  }

  return (
    <div className="turno-completo">
      <PanelTitle title="Turno completo" subtitle="Datos generales una sola vez" />
      <div className="capture-layout-tabs" aria-label="Modo de captura">
        <button className={captureLayout === 'asistente' ? 'active' : ''} type="button" onClick={() => setCaptureLayout('asistente')}>
          <ClipboardCheck size={17} /> Asistente
        </button>
        <button className={captureLayout === 'formulario' ? 'active' : ''} type="button" onClick={() => setCaptureLayout('formulario')}>
          <FileSpreadsheet size={17} /> Formulario completo
        </button>
      </div>
      <QuickSection title="Turno" icon={<Mountain size={18} />} anchorId="capture-turno">
        <ShiftBaseFields base={base} setBase={setBase} />
      </QuickSection>

      <QuickSection title="Trabajos del turno" icon={<ClipboardCheck size={18} />} anchorId="capture-modulos">
        <div className="quick-choice-grid three">
          <ChoiceButton active={modules.barrenacion} icon={<Drill size={21} />} label="Barrenacion / voladura" meta="PDF barrenacion" onClick={() => toggleModule('barrenacion')} />
          <ChoiceButton active={modules.rezagado} icon={<Truck size={21} />} label="Rezagado" meta="PDF rezagado" onClick={() => toggleModule('rezagado')} />
          <ChoiceButton active={modules.seguridad} icon={<ShieldCheck size={21} />} label="Seguridad" meta="Reporte local" onClick={() => toggleModule('seguridad')} />
        </div>
        <div className="turno-summary-strip">
          <article>
            <span>Modulos</span>
            <strong>{selectedCount}</strong>
          </article>
          <article>
            <span>Fecha</span>
            <strong>{base.fecha}</strong>
          </article>
          <article>
            <span>Turno</span>
            <strong>{base.turno}</strong>
          </article>
        </div>
      </QuickSection>

      {captureLayout === 'asistente' ? (
        <QuickSection title="Captura guiada" icon={<Pickaxe size={18} />} anchorId="capture-guiada">
          <div className="assistant-module-nav">
            {enabledModules.map((type) => (
              <button
                className={activeAssistantModule === type ? 'active' : ''}
                key={type}
                type="button"
                onClick={() => setAssistantModule(type)}
              >
                {type === 'barrenacion' && <Drill size={18} />}
                {type === 'rezagado' && <Truck size={18} />}
                {type === 'seguridad' && <ShieldCheck size={18} />}
                {shortTypeLabel(type)}
              </button>
            ))}
          </div>
          <div className="assistant-progress-grid">
            <AssistantMetric label="Jumbos" value={barrenacion.jumbo.length} />
            <AssistantMetric label="Voladuras" value={barrenacion.voladuras.length} />
            <AssistantMetric label="Scoop" value={rezagado.scoopTram.length} />
            <AssistantMetric label="Retro" value={rezagado.retro.length} />
          </div>
          <section className="turno-module-section assistant-active-module">
            {activeAssistantModule === 'barrenacion' && <BarrenacionAssistantForm equipmentFavorites={equipmentFavorites} record={barrenacion} setRecord={setBarrenacion} />}
            {activeAssistantModule === 'rezagado' && <RezagadoAssistantForm equipmentFavorites={equipmentFavorites} record={rezagado} setRecord={setRezagado} />}
            {activeAssistantModule === 'seguridad' && <SeguridadForm record={seguridad} setRecord={setSeguridad} showBaseFields={false} />}
          </section>
        </QuickSection>
      ) : (
        <>
          {modules.barrenacion && (
            <section className="turno-module-section">
              <BarrenacionForm record={barrenacion} setRecord={setBarrenacion} showBaseFields={false} />
            </section>
          )}
          {modules.rezagado && (
            <section className="turno-module-section">
              <RezagadoForm record={rezagado} setRecord={setRezagado} showBaseFields={false} />
            </section>
          )}
          {modules.seguridad && (
            <section className="turno-module-section">
              <SeguridadForm record={seguridad} setRecord={setSeguridad} showBaseFields={false} />
            </section>
          )}
        </>
      )}
    </div>
  )
}

function AssistantMetric({ label, value }: { label: string; value: number }) {
  return (
    <article>
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  )
}

function BarrenacionAssistantForm({
  equipmentFavorites,
  record,
  setRecord,
}: {
  equipmentFavorites: EquipmentFavorites
  record: BarrenacionRecord
  setRecord: (record: BarrenacionRecord) => void
}) {
  const [step, setStep] = useState<BarrenacionAssistantStep>('actividad')
  const [rowIndex, setRowIndex] = useState(0)
  const activity = getBarrenacionActivity(record)
  const steps: { key: BarrenacionAssistantStep; label: string }[] = [
    { key: 'actividad', label: 'Actividad' },
    { key: 'equipo', label: activity === 'voladura' ? 'Frente' : 'Equipo' },
    { key: 'produccion', label: 'Produccion' },
    { key: 'detalles', label: 'Cierre' },
  ]
  const isBlast = activity === 'voladura'
  const drillKind = activity === 'maquinaPierna' ? 'maquinaPierna' : 'jumbo'
  const jumboRows = record.jumbo.length ? record.jumbo : [emptyDrillRow(defaultJumbo)]
  const maquinaRows = record.maquinaPierna.length ? record.maquinaPierna : [emptyDrillRow(defaultJumbo)]
  const drillRows = drillKind === 'jumbo' ? jumboRows : maquinaRows
  const blastRows = record.voladuras.length ? record.voladuras : [emptyBlastRow()]
  const rowCount = isBlast ? blastRows.length : drillRows.length
  const safeIndex = Math.min(rowIndex, Math.max(rowCount - 1, 0))
  const drillRow = drillRows[safeIndex] ?? emptyDrillRow(defaultJumbo)
  const blastRow = blastRows[safeIndex] ?? emptyBlastRow()
  const drillEquipmentOptions = prioritizeOptions(
    jumboEquipoOptions,
    drillRow.equipo,
    [...equipmentFavorites.jumbo, ...drillRows.map((row) => row.equipo)],
  )

  function selectActivity(next: BarrenacionActivity) {
    setRowIndex(0)
    setStep('equipo')
    setRecord({
      ...record,
      activeActivity: next,
      jumbo: next === 'jumbo' && record.jumbo.length === 0 ? [emptyDrillRow(defaultJumbo)] : record.jumbo,
      maquinaPierna: next === 'maquinaPierna' && record.maquinaPierna.length === 0 ? [emptyDrillRow(defaultJumbo)] : record.maquinaPierna,
      voladuras: next === 'voladura' && record.voladuras.length === 0 ? [emptyBlastRow()] : record.voladuras,
    })
  }

  function setDrillRows(rows: DrillRow[]) {
    setRecord({
      ...record,
      activeActivity: drillKind,
      jumbo: drillKind === 'jumbo' ? rows : record.jumbo,
      maquinaPierna: drillKind === 'maquinaPierna' ? rows : record.maquinaPierna,
    })
  }

  function updateDrill(patch: Partial<DrillRow>) {
    setDrillRows(drillRows.map((row, index) => (index === safeIndex ? { ...row, ...patch } : row)))
  }

  function updateBlast(patch: Partial<BlastRow>) {
    setRecord({
      ...record,
      activeActivity: 'voladura',
      voladuras: blastRows.map((row, index) => (index === safeIndex ? { ...row, ...patch } : row)),
    })
  }

  function addCurrentRow() {
    if (isBlast) {
      setRecord({ ...record, activeActivity: 'voladura', voladuras: [...blastRows, emptyBlastRow()] })
      setRowIndex(blastRows.length)
      return
    }
    setDrillRows([...drillRows, emptyDrillRow(defaultJumbo)])
    setRowIndex(drillRows.length)
  }

  function duplicateCurrentRow() {
    if (isBlast) {
      const clone = { ...blastRow }
      const nextRows = [...blastRows.slice(0, safeIndex + 1), clone, ...blastRows.slice(safeIndex + 1)]
      setRecord({ ...record, activeActivity: 'voladura', voladuras: nextRows })
      setRowIndex(safeIndex + 1)
      return
    }
    const clone = { ...drillRow }
    const nextRows = [...drillRows.slice(0, safeIndex + 1), clone, ...drillRows.slice(safeIndex + 1)]
    setDrillRows(nextRows)
    setRowIndex(safeIndex + 1)
  }

  function removeCurrentRow() {
    if (isBlast) {
      const nextRows = blastRows.filter((_, index) => index !== safeIndex)
      setRecord({ ...record, activeActivity: 'voladura', voladuras: nextRows.length ? nextRows : [emptyBlastRow()] })
    } else {
      const nextRows = drillRows.filter((_, index) => index !== safeIndex)
      setDrillRows(nextRows.length ? nextRows : [emptyDrillRow(defaultJumbo)])
    }
    setRowIndex(Math.max(0, safeIndex - 1))
  }

  return (
    <div className="capture-assistant-shell">
      <PanelTitle title="Asistente barrenacion y voladuras" subtitle="Captura un registro corto y agrega mas equipos si hace falta" />
      <AssistantStepNav steps={steps} active={step} onChange={setStep} />

      <section className="assistant-focus-card">
        {step === 'actividad' && (
          <>
            <AssistantFocusTitle title="Selecciona el trabajo" meta="El asistente mostrara solo los campos necesarios para ese trabajo." />
            <div className="quick-choice-grid three">
              <ChoiceButton active={activity === 'jumbo'} icon={<Drill size={21} />} label="Jumbo" meta={`${record.jumbo.length || 1} registros`} onClick={() => selectActivity('jumbo')} />
              <ChoiceButton active={activity === 'maquinaPierna'} icon={<HardHat size={21} />} label="Maquina pierna" meta={`${record.maquinaPierna.length || 1} registros`} onClick={() => selectActivity('maquinaPierna')} />
              <ChoiceButton active={activity === 'voladura'} icon={<Activity size={21} />} label="Voladura" meta={`${record.voladuras.length || 1} registros`} onClick={() => selectActivity('voladura')} />
            </div>
          </>
        )}

        {step === 'equipo' && (
          <>
            <AssistantFocusTitle title={isBlast ? 'Frente de voladura' : 'Equipo y personal'} meta={`${shortBarrenacionActivityLabel(activity)} ${safeIndex + 1} de ${rowCount}`} />
            <AssistantRowTools
              addLabel={isBlast ? 'Agregar voladura' : 'Agregar equipo'}
              count={rowCount}
              index={safeIndex}
              label={isBlast ? 'Voladura' : shortBarrenacionActivityLabel(activity)}
              onAdd={addCurrentRow}
              onDuplicate={duplicateCurrentRow}
              onRemove={removeCurrentRow}
              onSelect={setRowIndex}
            />
            {isBlast ? (
              <div className="form-grid quick">
                <Field label="Obra / frente" value={blastRow.obra} onChange={(value) => updateBlast({ obra: value })} />
                <Field label="Oficial" value={blastRow.oficial} onChange={(value) => updateBlast({ oficial: value })} />
                <Field label="Ayudante" value={blastRow.ayudante} onChange={(value) => updateBlast({ ayudante: value })} />
                <Field label="RPA / Cfte" value={blastRow.rpaCfte} onChange={(value) => updateBlast({ rpaCfte: value })} />
              </div>
            ) : (
              <div className="form-grid quick">
                <Field label="Equipo" value={drillRow.equipo} options={drillEquipmentOptions} onChange={(value) => updateDrill({ equipo: value })} />
                <Field label="Operador" value={drillRow.operador} onChange={(value) => updateDrill({ operador: value })} />
                <Field label="Ayudante" value={drillRow.ayudante} onChange={(value) => updateDrill({ ayudante: value })} />
                <Field label="Nivel / obra" value={drillRow.nivelObra} onChange={(value) => updateDrill({ nivelObra: value })} />
                <Field label="RPA / Cfte" value={drillRow.rpaCfte} onChange={(value) => updateDrill({ rpaCfte: value })} />
              </div>
            )}
          </>
        )}

        {step === 'produccion' && (
          <>
            <AssistantFocusTitle title="Produccion del turno" meta={isBlast ? 'Barrenos y metros pegados' : 'Barrenos, metros y horas'} />
            {isBlast ? (
              <div className="form-grid quick">
                <Field label="Barrenos pegados" type="number" value={blastRow.barrenosPegados} onChange={(value) => updateBlast({ barrenosPegados: Number(value) })} />
                <Field label="Metros pegados" type="number" value={blastRow.metrosPegados} onChange={(value) => updateBlast({ metrosPegados: Number(value) })} />
                <Field label="Horas servicio" type="number" value={blastRow.horasServicio} onChange={(value) => updateBlast({ horasServicio: Number(value) })} />
                <Field label="Longitud" type="number" value={blastRow.longitud} onChange={(value) => updateBlast({ longitud: Number(value) })} />
                <Field label="Cuele" type="number" value={blastRow.cuele} onChange={(value) => updateBlast({ cuele: Number(value) })} />
                <Field label="Desarrollo" type="number" value={blastRow.desarrollo} onChange={(value) => updateBlast({ desarrollo: Number(value) })} />
              </div>
            ) : (
              <div className="form-grid quick">
                <Field label="Barrenos dados" type="number" value={drillRow.barrenosDados} onChange={(value) => updateDrill({ barrenosDados: Number(value) })} />
                <Field label="Barrenos cargados" type="number" value={drillRow.barrenosCargados} onChange={(value) => updateDrill({ barrenosCargados: Number(value) })} />
                <Field label="Metros dados" type="number" value={drillRow.metrosDados} onChange={(value) => updateDrill({ metrosDados: Number(value) })} />
                <Field label="Horas servicio" type="number" value={drillRow.horasServicio} onChange={(value) => updateDrill({ horasServicio: Number(value) })} />
                <Field label="Longitud" type="number" value={drillRow.longitud} onChange={(value) => updateDrill({ longitud: Number(value) })} />
                <Field label="Desarrollo" type="number" value={drillRow.desarrollo} onChange={(value) => updateDrill({ desarrollo: Number(value) })} />
              </div>
            )}
          </>
        )}

        {step === 'detalles' && (
          <>
            <AssistantFocusTitle title="Cierre y datos opcionales" meta="Captura solo lo que aplique; lo demas sigue disponible en Formulario completo." />
            {isBlast ? (
              <>
                <div className="form-grid quick">
                  <Field label="ANFO inicial" type="number" value={blastRow.anfoInicial} onChange={(value) => updateBlast({ anfoInicial: Number(value) })} />
                  <Field label="ANFO final" type="number" value={blastRow.anfoFinal} onChange={(value) => updateBlast({ anfoFinal: Number(value) })} />
                </div>
                <TextArea label="Observaciones de voladura" value={blastRow.observaciones} onChange={(value) => updateBlast({ observaciones: value })} />
              </>
            ) : (
              <details className="quick-details assistant-details">
                <summary>Horometros e insumos</summary>
                <div className="form-grid quick">
                  <Field label="Hor. diesel inicial" type="number" value={drillRow.horometroDieselInicial} onChange={(value) => updateDrill({ horometroDieselInicial: Number(value) })} />
                  <Field label="Hor. diesel final" type="number" value={drillRow.horometroDieselFinal} onChange={(value) => updateDrill({ horometroDieselFinal: Number(value) })} />
                  <Field label="Hor. elect inicial" type="number" value={drillRow.horometroElectInicial} onChange={(value) => updateDrill({ horometroElectInicial: Number(value) })} />
                  <Field label="Hor. elect final" type="number" value={drillRow.horometroElectFinal} onChange={(value) => updateDrill({ horometroElectFinal: Number(value) })} />
                  <Field label="Zanco" type="number" value={drillRow.zanco} onChange={(value) => updateDrill({ zanco: Number(value) })} />
                  <Field label="Cople" type="number" value={drillRow.cople} onChange={(value) => updateDrill({ cople: Number(value) })} />
                  <Field label="Barra" type="number" value={drillRow.barra} onChange={(value) => updateDrill({ barra: Number(value) })} />
                  <Field label="Broca" type="number" value={drillRow.broca} onChange={(value) => updateDrill({ broca: Number(value) })} />
                </div>
              </details>
            )}
            <details className="quick-details assistant-details">
              <summary>Apoyo del turno</summary>
              <div className="form-grid quick">
                <Field label="Polvorero" value={record.polvorero} onChange={(value) => setRecord({ ...record, polvorero: value })} />
                <Field label="Chofer camion personal" value={record.choferCamion} onChange={(value) => setRecord({ ...record, choferCamion: value })} />
                <Field label="Chofer pipa" value={record.choferPipa} onChange={(value) => setRecord({ ...record, choferPipa: value })} />
                <Field label="Bob cat" value={record.bobCat} onChange={(value) => setRecord({ ...record, bobCat: value })} />
                <Field label="Bombeo" value={record.bombeo} onChange={(value) => setRecord({ ...record, bombeo: value })} />
                <Field label="Servicios" value={record.servicios} onChange={(value) => setRecord({ ...record, servicios: value })} />
              </div>
            </details>
            <TextArea label="Comentarios generales" value={record.comentarios} onChange={(value) => setRecord({ ...record, comentarios: value })} />
            <TextArea label="Inasistencias / permisos" value={record.inasistencias} onChange={(value) => setRecord({ ...record, inasistencias: value })} />
          </>
        )}
      </section>

      <AssistantStepActions steps={steps} active={step} onChange={setStep} />
    </div>
  )
}

function RezagadoAssistantForm({
  equipmentFavorites,
  record,
  setRecord,
}: {
  equipmentFavorites: EquipmentFavorites
  record: RezagadoRecord
  setRecord: (record: RezagadoRecord) => void
}) {
  const [step, setStep] = useState<RezagadoAssistantStep>('equipo')
  const [rowIndex, setRowIndex] = useState(0)
  const equipment = getRezagadoEquipment(record)
  const steps: { key: RezagadoAssistantStep; label: string }[] = [
    { key: 'equipo', label: 'Equipo' },
    { key: 'trabajo', label: 'Trabajo' },
    { key: 'produccion', label: 'Produccion' },
    { key: 'cierre', label: 'Cierre' },
  ]
  const isScoop = equipment === 'scoop'
  const scoopRows = record.scoopTram.length ? record.scoopTram : [emptyHaulRow(defaultScoop)]
  const retroRows = record.retro.length ? record.retro : [emptyRetroRow()]
  const rowCount = isScoop ? scoopRows.length : retroRows.length
  const safeIndex = Math.min(rowIndex, Math.max(rowCount - 1, 0))
  const scoopRow = scoopRows[safeIndex] ?? emptyHaulRow(defaultScoop)
  const retroRow = retroRows[safeIndex] ?? emptyRetroRow()
  const selectedScoopActivities = getSelectedHaulActivities(scoopRow)
  const selectedRetroActivities = getSelectedRetroActivities(retroRow)
  const scoopOptions = prioritizeOptions(
    scoopEquipoOptions,
    scoopRow.equipo,
    [...equipmentFavorites.scoop, ...scoopRows.map((row) => row.equipo)],
  )
  const retroOptions = prioritizeOptions(
    retroEquipoOptions,
    retroRow.equipo,
    [...equipmentFavorites.retro, ...retroRows.map((row) => row.equipo)],
  )

  function selectEquipment(next: RezagadoEquipment) {
    setRowIndex(0)
    setStep('equipo')
    setRecord({
      ...record,
      activeEquipment: next,
      scoopTram: next === 'scoop' && record.scoopTram.length === 0 ? [emptyHaulRow(defaultScoop)] : record.scoopTram,
      retro: next === 'retro' && record.retro.length === 0 ? [emptyRetroRow()] : record.retro,
    })
  }

  function setScoopRows(rows: HaulRow[]) {
    setRecord({ ...record, activeEquipment: 'scoop', scoopTram: rows.length ? rows : [emptyHaulRow(defaultScoop)] })
  }

  function setRetroRows(rows: RetroRow[]) {
    setRecord({ ...record, activeEquipment: 'retro', retro: rows.length ? rows : [emptyRetroRow()] })
  }

  function updateScoopRow(patch: Partial<HaulRow>) {
    setScoopRows(scoopRows.map((row, index) => (index === safeIndex ? { ...row, ...patch } : row)))
  }

  function updateRetroRow(patch: Partial<RetroRow>) {
    setRetroRows(retroRows.map((row, index) => (index === safeIndex ? { ...row, ...patch } : row)))
  }

  function toggleScoopActivity(key: HaulActivity) {
    const nextSelected = selectedScoopActivities.includes(key) ? selectedScoopActivities.filter((item) => item !== key) : [...selectedScoopActivities, key]
    if (nextSelected.length === 0) return
    updateScoopRow({ selectedActivities: nextSelected, [key]: nextSelected.includes(key) ? scoopRow[key] : 0 } as Partial<HaulRow>)
  }

  function toggleRetroActivity(key: RetroActivity) {
    const nextSelected = selectedRetroActivities.includes(key) ? selectedRetroActivities.filter((item) => item !== key) : [...selectedRetroActivities, key]
    if (nextSelected.length === 0) return
    updateRetroRow({ selectedActivities: nextSelected, [key]: nextSelected.includes(key) ? retroRow[key] : 0 } as Partial<RetroRow>)
  }

  function addCurrentRow() {
    if (isScoop) {
      setScoopRows([...scoopRows, emptyHaulRow(defaultScoop)])
      setRowIndex(scoopRows.length)
      return
    }
    setRetroRows([...retroRows, emptyRetroRow()])
    setRowIndex(retroRows.length)
  }

  function duplicateCurrentRow() {
    if (isScoop) {
      const clone = { ...scoopRow, selectedActivities: [...selectedScoopActivities] }
      const nextRows = [...scoopRows.slice(0, safeIndex + 1), clone, ...scoopRows.slice(safeIndex + 1)]
      setScoopRows(nextRows)
      setRowIndex(safeIndex + 1)
      return
    }
    const clone = { ...retroRow, selectedActivities: [...selectedRetroActivities] }
    const nextRows = [...retroRows.slice(0, safeIndex + 1), clone, ...retroRows.slice(safeIndex + 1)]
    setRetroRows(nextRows)
    setRowIndex(safeIndex + 1)
  }

  function removeCurrentRow() {
    if (isScoop) {
      const nextRows = scoopRows.filter((_, index) => index !== safeIndex)
      setScoopRows(nextRows.length ? nextRows : [emptyHaulRow(defaultScoop)])
    } else {
      const nextRows = retroRows.filter((_, index) => index !== safeIndex)
      setRetroRows(nextRows.length ? nextRows : [emptyRetroRow()])
    }
    setRowIndex(Math.max(0, safeIndex - 1))
  }

  return (
    <div className="capture-assistant-shell">
      <PanelTitle title="Asistente rezagado" subtitle="Selecciona actividades y captura cantidades por equipo" />
      <AssistantStepNav steps={steps} active={step} onChange={setStep} />

      <section className="assistant-focus-card">
        {step === 'equipo' && (
          <>
            <AssistantFocusTitle title="Equipo de rezagado" meta={`${isScoop ? 'Scoop tram' : 'Retro'} ${safeIndex + 1} de ${rowCount}`} />
            <div className="quick-choice-grid two">
              <ChoiceButton active={equipment === 'scoop'} icon={<Truck size={21} />} label="Scoop tram" meta={`${record.scoopTram.length || 1} registros`} onClick={() => selectEquipment('scoop')} />
              <ChoiceButton active={equipment === 'retro'} icon={<HardHat size={21} />} label="Retro" meta={`${record.retro.length || 1} registros`} onClick={() => selectEquipment('retro')} />
            </div>
            <AssistantRowTools
              addLabel="Agregar equipo"
              count={rowCount}
              index={safeIndex}
              label={isScoop ? 'Scoop' : 'Retro'}
              onAdd={addCurrentRow}
              onDuplicate={duplicateCurrentRow}
              onRemove={removeCurrentRow}
              onSelect={setRowIndex}
            />
            {isScoop ? (
              <div className="form-grid quick">
                <Field label="Equipo" value={scoopRow.equipo} options={scoopOptions} onChange={(value) => updateScoopRow({ equipo: value })} />
                <Field label="Operador" value={scoopRow.operador} onChange={(value) => updateScoopRow({ operador: value })} />
                <Field label="Nivel / obra" value={scoopRow.nivelObra} onChange={(value) => updateScoopRow({ nivelObra: value })} />
                <Field label="Destino" value={scoopRow.destino} onChange={(value) => updateScoopRow({ destino: value })} />
              </div>
            ) : (
              <div className="form-grid quick">
                <Field label="Equipo" value={retroRow.equipo} options={retroOptions} onChange={(value) => updateRetroRow({ equipo: value })} />
                <Field label="Operador" value={retroRow.operador} onChange={(value) => updateRetroRow({ operador: value })} />
                <Field label="Nivel / obra" value={retroRow.nivelObra} onChange={(value) => updateRetroRow({ nivelObra: value })} />
              </div>
            )}
          </>
        )}

        {step === 'trabajo' && (
          <>
            <AssistantFocusTitle title="Trabajo realizado" meta="Puedes seleccionar varios trabajos para el mismo equipo." />
            <div className="quick-choice-grid compact multi-choice-grid assistant-work-grid">
              {isScoop
                ? haulActivityKeys.map((key) => (
                  <ChoiceButton
                    key={key}
                    active={selectedScoopActivities.includes(key)}
                    label={haulActivityLabels[key]}
                    onClick={() => toggleScoopActivity(key)}
                  />
                ))
                : retroActivityKeys.map((key) => (
                  <ChoiceButton
                    key={key}
                    active={selectedRetroActivities.includes(key)}
                    label={retroActivityLabels[key]}
                    onClick={() => toggleRetroActivity(key)}
                  />
                ))}
            </div>
          </>
        )}

        {step === 'produccion' && (
          <>
            <AssistantFocusTitle title="Cantidades y consumo" meta="Solo aparecen las actividades seleccionadas." />
            {isScoop ? (
              <div className="form-grid quick">
                {selectedScoopActivities.map((key) => (
                  <Field key={key} label={`Cantidad ${haulActivityLabels[key]}`} type="number" value={scoopRow[key]} onChange={(value) => updateScoopRow({ [key]: Number(value) } as Partial<HaulRow>)} />
                ))}
                <Field label="Camiones" type="number" value={scoopRow.camiones} onChange={(value) => updateScoopRow({ camiones: Number(value) })} />
                <Field label="Hor. inicial" type="number" value={scoopRow.horometroInicial} onChange={(value) => updateScoopRow({ horometroInicial: Number(value) })} />
                <Field label="Hor. final" type="number" value={scoopRow.horometroFinal} onChange={(value) => updateScoopRow({ horometroFinal: Number(value) })} />
                <Field label="Diesel" type="number" value={scoopRow.diesel} onChange={(value) => updateScoopRow({ diesel: Number(value) })} />
              </div>
            ) : (
              <div className="form-grid quick">
                {selectedRetroActivities.map((key) => (
                  <Field key={key} label={`Cantidad ${retroActivityLabels[key]}`} type="number" value={retroRow[key]} onChange={(value) => updateRetroRow({ [key]: Number(value) } as Partial<RetroRow>)} />
                ))}
                <Field label="Hor. inicial" type="number" value={retroRow.horometroInicial} onChange={(value) => updateRetroRow({ horometroInicial: Number(value) })} />
                <Field label="Hor. final" type="number" value={retroRow.horometroFinal} onChange={(value) => updateRetroRow({ horometroFinal: Number(value) })} />
                <Field label="Diesel" type="number" value={retroRow.diesel} onChange={(value) => updateRetroRow({ diesel: Number(value) })} />
              </div>
            )}
          </>
        )}

        {step === 'cierre' && (
          <>
            <AssistantFocusTitle title="Observaciones del turno" meta="El detalle queda guardado en historial aunque no sincronices todavia." />
            {isScoop ? (
              <TextArea label="Observaciones del equipo" value={scoopRow.observaciones} onChange={(value) => updateScoopRow({ observaciones: value })} />
            ) : (
              <TextArea label="Observaciones del equipo" value={retroRow.observaciones} onChange={(value) => updateRetroRow({ observaciones: value })} />
            )}
            <TextArea label="Comentarios generales" value={record.comentarios} onChange={(value) => setRecord({ ...record, comentarios: value })} />
          </>
        )}
      </section>

      <AssistantStepActions steps={steps} active={step} onChange={setStep} />
    </div>
  )
}

function AssistantStepNav<T extends string>({
  active,
  onChange,
  steps,
}: {
  active: T
  onChange: (step: T) => void
  steps: { key: T; label: string }[]
}) {
  return (
    <div className="assistant-stepper" aria-label="Pasos de captura">
      {steps.map((item, index) => (
        <button className={item.key === active ? 'active' : ''} key={item.key} type="button" onClick={() => onChange(item.key)}>
          <span>{index + 1}</span>
          <strong>{item.label}</strong>
        </button>
      ))}
    </div>
  )
}

function AssistantStepActions<T extends string>({
  active,
  onChange,
  steps,
}: {
  active: T
  onChange: (step: T) => void
  steps: { key: T; label: string }[]
}) {
  const index = Math.max(0, steps.findIndex((item) => item.key === active))
  const previous = steps[Math.max(0, index - 1)]
  const next = steps[Math.min(steps.length - 1, index + 1)]
  return (
    <div className="assistant-footer-actions">
      <button type="button" onClick={() => onChange(previous.key)} disabled={index === 0}>
        Anterior
      </button>
      <button className="assistant-next-action" type="button" onClick={() => onChange(next.key)} disabled={index === steps.length - 1}>
        Siguiente
      </button>
    </div>
  )
}

function AssistantFocusTitle({ title, meta }: { title: string; meta: string }) {
  return (
    <div className="assistant-focus-title">
      <div>
        <span>{meta}</span>
        <h2>{title}</h2>
      </div>
    </div>
  )
}

function AssistantRowTools({
  addLabel,
  count,
  index,
  label,
  onAdd,
  onDuplicate,
  onRemove,
  onSelect,
}: {
  addLabel: string
  count: number
  index: number
  label: string
  onAdd: () => void
  onDuplicate: () => void
  onRemove: () => void
  onSelect: (index: number) => void
}) {
  return (
    <div className="assistant-row-tools">
      <div className="assistant-row-tabs" aria-label="Registros capturados">
        {Array.from({ length: count }, (_, rowIndex) => (
          <button className={rowIndex === index ? 'active' : ''} key={rowIndex} type="button" onClick={() => onSelect(rowIndex)}>
            {label} {rowIndex + 1}
          </button>
        ))}
      </div>
      <button className="assistant-small-action" type="button" onClick={onAdd}>
        <Plus size={15} /> {addLabel}
      </button>
      <button className="assistant-small-action" type="button" onClick={onDuplicate}>
        <Edit3 size={15} /> Duplicar
      </button>
      <button className="assistant-small-action danger" type="button" onClick={onRemove}>
        <Trash2 size={15} /> Eliminar
      </button>
    </div>
  )
}

function ShiftBaseFields({ base, setBase }: { base: ShiftBase; setBase: (base: ShiftBase) => void }) {
  return (
    <div className="form-grid four">
      <Field label="Supervisor" value={base.supervisor} onChange={(value) => setBase({ ...base, supervisor: value })} required />
      <Field label="Fecha" type="date" value={base.fecha} onChange={(value) => setBase({ ...base, fecha: value })} required />
      <label>
        Turno
        <select value={base.turno} onChange={(event) => setBase({ ...base, turno: event.target.value as Shift })}>
          <option value="1">Turno 1</option>
          <option value="2">Turno 2</option>
        </select>
      </label>
      <Field label="Unidad" value={base.unidad} onChange={(value) => setBase({ ...base, unidad: value })} />
    </div>
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

function BarrenacionForm({
  record,
  setRecord,
  showBaseFields = true,
}: {
  record: BarrenacionRecord
  setRecord: (record: BarrenacionRecord) => void
  showBaseFields?: boolean
}) {
  const activity = getBarrenacionActivity(record)
  const jumboRows = record.jumbo.length ? record.jumbo : [emptyDrillRow(defaultJumbo)]
  const maquinaRows = record.maquinaPierna.length ? record.maquinaPierna : [emptyDrillRow(defaultJumbo)]
  const blastRows = record.voladuras.length ? record.voladuras : [emptyBlastRow()]
  const activeDrillRows = activity === 'maquinaPierna' ? maquinaRows : jumboRows

  function selectActivity(next: BarrenacionActivity) {
    setRecord({
      ...record,
      activeActivity: next,
      jumbo: next === 'jumbo' && record.jumbo.length === 0 ? [emptyDrillRow(defaultJumbo)] : record.jumbo,
      maquinaPierna: next === 'maquinaPierna' && record.maquinaPierna.length === 0 ? [emptyDrillRow(defaultJumbo)] : record.maquinaPierna,
      voladuras: next === 'voladura' && record.voladuras.length === 0 ? [emptyBlastRow()] : record.voladuras,
    })
  }

  function setDrillRows(kind: 'jumbo' | 'maquinaPierna', rows: DrillRow[]) {
    setRecord({
      ...record,
      activeActivity: kind,
      jumbo: kind === 'jumbo' ? rows : record.jumbo,
      maquinaPierna: kind === 'maquinaPierna' ? rows : record.maquinaPierna,
    })
  }

  function updateDrill(kind: 'jumbo' | 'maquinaPierna', index: number, patch: Partial<DrillRow>) {
    const rows = kind === 'jumbo' ? jumboRows : maquinaRows
    setDrillRows(kind, rows.map((row, rowIndex) => (rowIndex === index ? { ...row, ...patch } : row)))
  }

  function addDrill(kind: 'jumbo' | 'maquinaPierna') {
    const rows = kind === 'jumbo' ? jumboRows : maquinaRows
    setDrillRows(kind, [...rows, emptyDrillRow(defaultJumbo)])
  }

  function removeDrill(kind: 'jumbo' | 'maquinaPierna', index: number) {
    const rows = kind === 'jumbo' ? jumboRows : maquinaRows
    const nextRows = rows.filter((_, rowIndex) => rowIndex !== index)
    setDrillRows(kind, nextRows.length ? nextRows : [emptyDrillRow(defaultJumbo)])
  }

  function updateBlast(index: number, patch: Partial<BlastRow>) {
    setRecord({
      ...record,
      activeActivity: 'voladura',
      voladuras: blastRows.map((row, rowIndex) => (rowIndex === index ? { ...row, ...patch } : row)),
    })
  }

  function addBlast() {
    setRecord({ ...record, activeActivity: 'voladura', voladuras: [...blastRows, emptyBlastRow()] })
  }

  function removeBlast(index: number) {
    const nextRows = blastRows.filter((_, rowIndex) => rowIndex !== index)
    setRecord({ ...record, activeActivity: 'voladura', voladuras: nextRows.length ? nextRows : [emptyBlastRow()] })
  }

  return (
    <>
      <PanelTitle title="Captura de turno barrenacion y voladuras" subtitle="Varios equipos por supervisor" />
      {showBaseFields && (
        <QuickSection title="Turno" icon={<Mountain size={18} />} anchorId="capture-turno">
          <BaseFields record={record} setRecord={setRecord} />
        </QuickSection>
      )}
      <QuickSection title="Actividad" icon={<Drill size={18} />} anchorId="capture-actividad">
        <div className="quick-choice-grid three">
          <ChoiceButton active={activity === 'jumbo'} icon={<Drill size={21} />} label="Jumbo" meta={`${record.jumbo.length || 1} filas`} onClick={() => selectActivity('jumbo')} />
          <ChoiceButton active={activity === 'maquinaPierna'} icon={<HardHat size={21} />} label="Maquina pierna" meta={`${record.maquinaPierna.length || 1} filas`} onClick={() => selectActivity('maquinaPierna')} />
          <ChoiceButton active={activity === 'voladura'} icon={<Activity size={21} />} label="Voladura" meta={`${record.voladuras.length || 1} filas`} onClick={() => selectActivity('voladura')} />
        </div>
      </QuickSection>

      {activity === 'voladura' ? (
        <QuickSection title="Voladuras del turno" icon={<Activity size={18} />} anchorId="capture-produccion">
          <div className="section-heading">
            <h2>Filas de voladura</h2>
            <button className="inline-action" type="button" onClick={addBlast}>
              <Plus size={16} /> Agregar voladura
            </button>
          </div>
          {blastRows.map((row, index) => (
            <article className="row-card" key={`blast-${index}`}>
              <div className="row-title">
                <div>
                  <strong>Voladura {index + 1}</strong>
                  <small>Se llena en el formato de barrenacion y voladuras</small>
                </div>
                <button className="icon-button danger" type="button" onClick={() => removeBlast(index)} aria-label="Eliminar voladura">
                  <Trash2 size={16} />
                </button>
              </div>
              <div className="form-grid quick">
                <Field label="Obra / frente" value={row.obra} onChange={(value) => updateBlast(index, { obra: value })} />
                <Field label="Oficial" value={row.oficial} onChange={(value) => updateBlast(index, { oficial: value })} />
                <Field label="Ayudante" value={row.ayudante} onChange={(value) => updateBlast(index, { ayudante: value })} />
                <Field label="RPA / Cfte" value={row.rpaCfte} onChange={(value) => updateBlast(index, { rpaCfte: value })} />
                <Field label="Barrenos pegados" type="number" value={row.barrenosPegados} onChange={(value) => updateBlast(index, { barrenosPegados: Number(value) })} />
                <Field label="Metros pegados" type="number" value={row.metrosPegados} onChange={(value) => updateBlast(index, { metrosPegados: Number(value) })} />
                <Field label="Horas servicio" type="number" value={row.horasServicio} onChange={(value) => updateBlast(index, { horasServicio: Number(value) })} />
              </div>
              <details className="quick-details">
                <summary>Insumos y avance</summary>
                <div className="form-grid quick">
                  <Field label="Longitud" type="number" value={row.longitud} onChange={(value) => updateBlast(index, { longitud: Number(value) })} />
                  <Field label="Cuele" type="number" value={row.cuele} onChange={(value) => updateBlast(index, { cuele: Number(value) })} />
                  <Field label="Desarrollo" type="number" value={row.desarrollo} onChange={(value) => updateBlast(index, { desarrollo: Number(value) })} />
                  <Field label="ANFO inicial" type="number" value={row.anfoInicial} onChange={(value) => updateBlast(index, { anfoInicial: Number(value) })} />
                  <Field label="ANFO final" type="number" value={row.anfoFinal} onChange={(value) => updateBlast(index, { anfoFinal: Number(value) })} />
                </div>
              </details>
              <TextArea label="Observaciones" value={row.observaciones} onChange={(value) => updateBlast(index, { observaciones: value })} />
            </article>
          ))}
        </QuickSection>
      ) : (
        <QuickSection title={activity === 'jumbo' ? 'Jumbos del turno' : 'Maquina pierna del turno'} icon={<Pickaxe size={18} />} anchorId="capture-produccion">
          <div className="section-heading">
            <h2>{activity === 'jumbo' ? 'Filas de jumbo' : 'Filas de maquina pierna'}</h2>
            <button className="inline-action" type="button" onClick={() => addDrill(activity)}>
              <Plus size={16} /> Agregar equipo
            </button>
          </div>
          {activeDrillRows.map((row, index) => (
            <article className="row-card" key={`${activity}-${index}`}>
              <div className="row-title">
                <div>
                  <strong>{activity === 'jumbo' ? 'Jumbo' : 'Maquina pierna'} {index + 1}</strong>
                  <small>Un equipo o frente trabajado durante el turno</small>
                </div>
                <button className="icon-button danger" type="button" onClick={() => removeDrill(activity, index)} aria-label="Eliminar equipo">
                  <Trash2 size={16} />
                </button>
              </div>
              <div className="form-grid quick">
                <Field label="Equipo" value={row.equipo} options={jumboEquipoOptions} onChange={(value) => updateDrill(activity, index, { equipo: value })} />
                <Field label="Operador" value={row.operador} onChange={(value) => updateDrill(activity, index, { operador: value })} />
                <Field label="Ayudante" value={row.ayudante} onChange={(value) => updateDrill(activity, index, { ayudante: value })} />
                <Field label="Nivel / obra" value={row.nivelObra} onChange={(value) => updateDrill(activity, index, { nivelObra: value })} />
                <Field label="RPA / Cfte" value={row.rpaCfte} onChange={(value) => updateDrill(activity, index, { rpaCfte: value })} />
                <Field label="Barrenos dados" type="number" value={row.barrenosDados} onChange={(value) => updateDrill(activity, index, { barrenosDados: Number(value) })} />
                <Field label="Barrenos cargados" type="number" value={row.barrenosCargados} onChange={(value) => updateDrill(activity, index, { barrenosCargados: Number(value) })} />
                <Field label="Metros dados" type="number" value={row.metrosDados} onChange={(value) => updateDrill(activity, index, { metrosDados: Number(value) })} />
                <Field label="Horas servicio" type="number" value={row.horasServicio} onChange={(value) => updateDrill(activity, index, { horasServicio: Number(value) })} />
              </div>
              <details className="quick-details">
                <summary>Horometros e insumos</summary>
                <div className="form-grid quick">
                  <Field label="Longitud" type="number" value={row.longitud} onChange={(value) => updateDrill(activity, index, { longitud: Number(value) })} />
                  <Field label="Cuele" type="number" value={row.cuele} onChange={(value) => updateDrill(activity, index, { cuele: Number(value) })} />
                  <Field label="Desarrollo" type="number" value={row.desarrollo} onChange={(value) => updateDrill(activity, index, { desarrollo: Number(value) })} />
                  <Field label="Hor. diesel inicial" type="number" value={row.horometroDieselInicial} onChange={(value) => updateDrill(activity, index, { horometroDieselInicial: Number(value) })} />
                  <Field label="Hor. diesel final" type="number" value={row.horometroDieselFinal} onChange={(value) => updateDrill(activity, index, { horometroDieselFinal: Number(value) })} />
                  <Field label="Hor. elect inicial" type="number" value={row.horometroElectInicial} onChange={(value) => updateDrill(activity, index, { horometroElectInicial: Number(value) })} />
                  <Field label="Hor. elect final" type="number" value={row.horometroElectFinal} onChange={(value) => updateDrill(activity, index, { horometroElectFinal: Number(value) })} />
                  <Field label="Zanco" type="number" value={row.zanco} onChange={(value) => updateDrill(activity, index, { zanco: Number(value) })} />
                  <Field label="Cople" type="number" value={row.cople} onChange={(value) => updateDrill(activity, index, { cople: Number(value) })} />
                  <Field label="Barra" type="number" value={row.barra} onChange={(value) => updateDrill(activity, index, { barra: Number(value) })} />
                  <Field label="Broca" type="number" value={row.broca} onChange={(value) => updateDrill(activity, index, { broca: Number(value) })} />
                </div>
              </details>
            </article>
          ))}
        </QuickSection>
      )}

      <QuickSection title="Apoyo y cierre del turno" icon={<ClipboardCheck size={18} />} anchorId="capture-cierre">
        <div className="form-grid quick">
          <Field label="Polvorero" value={record.polvorero} onChange={(value) => setRecord({ ...record, polvorero: value })} />
          <Field label="Chofer camion personal" value={record.choferCamion} onChange={(value) => setRecord({ ...record, choferCamion: value })} />
          <Field label="Chofer pipa" value={record.choferPipa} onChange={(value) => setRecord({ ...record, choferPipa: value })} />
          <Field label="Bob cat" value={record.bobCat} onChange={(value) => setRecord({ ...record, bobCat: value })} />
          <Field label="Bombeo" value={record.bombeo} onChange={(value) => setRecord({ ...record, bombeo: value })} />
          <Field label="Servicios" value={record.servicios} onChange={(value) => setRecord({ ...record, servicios: value })} />
        </div>
        <TextArea label="Comentarios generales" value={record.comentarios} onChange={(value) => setRecord({ ...record, comentarios: value })} />
        <TextArea label="Inasistencias / permisos" value={record.inasistencias} onChange={(value) => setRecord({ ...record, inasistencias: value })} />
      </QuickSection>
    </>
  )
}

function RezagadoForm({
  record,
  setRecord,
  showBaseFields = true,
}: {
  record: RezagadoRecord
  setRecord: (record: RezagadoRecord) => void
  showBaseFields?: boolean
}) {
  const equipment = getRezagadoEquipment(record)
  const scoopRows = record.scoopTram.length ? record.scoopTram : [emptyHaulRow(defaultScoop)]
  const retroRows = record.retro.length ? record.retro : [emptyRetroRow()]

  function selectEquipment(next: RezagadoEquipment) {
    setRecord({
      ...record,
      activeEquipment: next,
      scoopTram: next === 'scoop' && record.scoopTram.length === 0 ? [emptyHaulRow(defaultScoop)] : record.scoopTram,
      retro: next === 'retro' && record.retro.length === 0 ? [emptyRetroRow()] : record.retro,
    })
  }

  function setScoopRows(rows: HaulRow[]) {
    setRecord({ ...record, activeEquipment: 'scoop', scoopTram: rows.length ? rows : [emptyHaulRow(defaultScoop)] })
  }

  function setRetroRows(rows: RetroRow[]) {
    setRecord({ ...record, activeEquipment: 'retro', retro: rows.length ? rows : [emptyRetroRow()] })
  }

  function updateScoopRow(index: number, patch: Partial<HaulRow>) {
    setScoopRows(scoopRows.map((row, rowIndex) => (rowIndex === index ? { ...row, ...patch } : row)))
  }

  function updateRetroRow(index: number, patch: Partial<RetroRow>) {
    setRetroRows(retroRows.map((row, rowIndex) => (rowIndex === index ? { ...row, ...patch } : row)))
  }

  function toggleScoopActivity(index: number, key: HaulActivity) {
    const row = scoopRows[index]
    const selected = getSelectedHaulActivities(row)
    const nextSelected = selected.includes(key) ? selected.filter((item) => item !== key) : [...selected, key]
    if (nextSelected.length === 0) return
    updateScoopRow(index, { selectedActivities: nextSelected, [key]: nextSelected.includes(key) ? row[key] : 0 } as Partial<HaulRow>)
  }

  function toggleRetroActivity(index: number, key: RetroActivity) {
    const row = retroRows[index]
    const selected = getSelectedRetroActivities(row)
    const nextSelected = selected.includes(key) ? selected.filter((item) => item !== key) : [...selected, key]
    if (nextSelected.length === 0) return
    updateRetroRow(index, { selectedActivities: nextSelected, [key]: nextSelected.includes(key) ? row[key] : 0 } as Partial<RetroRow>)
  }

  return (
    <>
      <PanelTitle title="Captura de turno rezagado" subtitle="Varios equipos y actividades" />
      {showBaseFields && (
        <QuickSection title="Turno" icon={<Mountain size={18} />} anchorId="capture-turno">
          <BaseFields record={record} setRecord={setRecord} />
        </QuickSection>
      )}
      <QuickSection title="Equipo" icon={<Truck size={18} />} anchorId="capture-equipo">
        <div className="quick-choice-grid two">
          <ChoiceButton active={equipment === 'scoop'} icon={<Truck size={21} />} label="Scoop tram" meta={`${record.scoopTram.length || 1} filas`} onClick={() => selectEquipment('scoop')} />
          <ChoiceButton active={equipment === 'retro'} icon={<HardHat size={21} />} label="Retro" meta={`${record.retro.length || 1} filas`} onClick={() => selectEquipment('retro')} />
        </div>
      </QuickSection>
      <QuickSection title="Trabajo" icon={<Activity size={18} />} anchorId="capture-trabajo">
        {equipment === 'scoop' ? (
          <>
            <div className="section-heading">
              <h2>Scoop tram del turno</h2>
              <button className="inline-action" type="button" onClick={() => setScoopRows([...scoopRows, emptyHaulRow(defaultScoop)])}>
                <Plus size={16} /> Agregar equipo
              </button>
            </div>
            {scoopRows.map((row, index) => {
              const selectedActivities = getSelectedHaulActivities(row)
              return (
                <article className="row-card" key={`scoop-${index}`}>
                  <div className="row-title">
                    <div>
                      <strong>Scoop tram {index + 1}</strong>
                      <small>Selecciona todas las actividades realizadas por este equipo</small>
                    </div>
                    <button className="icon-button danger" type="button" onClick={() => setScoopRows(scoopRows.filter((_, rowIndex) => rowIndex !== index))} aria-label="Eliminar scoop">
                      <Trash2 size={16} />
                    </button>
                  </div>
                  <div className="form-grid quick">
                    <Field label="Equipo" value={row.equipo} options={scoopEquipoOptions} onChange={(value) => updateScoopRow(index, { equipo: value })} />
                    <Field label="Operador" value={row.operador} onChange={(value) => updateScoopRow(index, { operador: value })} />
                    <Field label="Nivel / obra" value={row.nivelObra} onChange={(value) => updateScoopRow(index, { nivelObra: value })} />
                    <Field label="Destino" value={row.destino} onChange={(value) => updateScoopRow(index, { destino: value })} />
                  </div>
                  <div className="quick-choice-grid compact multi-choice-grid">
                    {haulActivityKeys.map((key) => (
                      <ChoiceButton
                        key={key}
                        active={selectedActivities.includes(key)}
                        label={haulActivityLabels[key]}
                        onClick={() => toggleScoopActivity(index, key)}
                      />
                    ))}
                  </div>
                  <div className="form-grid quick">
                    {selectedActivities.map((key) => (
                      <Field key={key} label={`Cantidad ${haulActivityLabels[key]}`} type="number" value={row[key]} onChange={(value) => updateScoopRow(index, { [key]: Number(value) } as Partial<HaulRow>)} />
                    ))}
                    <Field label="Camiones" type="number" value={row.camiones} onChange={(value) => updateScoopRow(index, { camiones: Number(value) })} />
                    <Field label="Hor. inicial" type="number" value={row.horometroInicial} onChange={(value) => updateScoopRow(index, { horometroInicial: Number(value) })} />
                    <Field label="Hor. final" type="number" value={row.horometroFinal} onChange={(value) => updateScoopRow(index, { horometroFinal: Number(value) })} />
                    <Field label="Diesel" type="number" value={row.diesel} onChange={(value) => updateScoopRow(index, { diesel: Number(value) })} />
                  </div>
                  <TextArea label="Observaciones" value={row.observaciones} onChange={(value) => updateScoopRow(index, { observaciones: value })} />
                </article>
              )
            })}
          </>
        ) : (
          <>
            <div className="section-heading">
              <h2>Retro del turno</h2>
              <button className="inline-action" type="button" onClick={() => setRetroRows([...retroRows, emptyRetroRow()])}>
                <Plus size={16} /> Agregar equipo
              </button>
            </div>
            {retroRows.map((row, index) => {
              const selectedActivities = getSelectedRetroActivities(row)
              return (
                <article className="row-card" key={`retro-${index}`}>
                  <div className="row-title">
                    <div>
                      <strong>Retro {index + 1}</strong>
                      <small>Selecciona todas las actividades realizadas por este equipo</small>
                    </div>
                    <button className="icon-button danger" type="button" onClick={() => setRetroRows(retroRows.filter((_, rowIndex) => rowIndex !== index))} aria-label="Eliminar retro">
                      <Trash2 size={16} />
                    </button>
                  </div>
                  <div className="form-grid quick">
                    <Field label="Equipo" value={row.equipo} options={retroEquipoOptions} onChange={(value) => updateRetroRow(index, { equipo: value })} />
                    <Field label="Operador" value={row.operador} onChange={(value) => updateRetroRow(index, { operador: value })} />
                    <Field label="Nivel / obra" value={row.nivelObra} onChange={(value) => updateRetroRow(index, { nivelObra: value })} />
                  </div>
                  <div className="quick-choice-grid compact multi-choice-grid">
                    {retroActivityKeys.map((key) => (
                      <ChoiceButton
                        key={key}
                        active={selectedActivities.includes(key)}
                        label={retroActivityLabels[key]}
                        onClick={() => toggleRetroActivity(index, key)}
                      />
                    ))}
                  </div>
                  <div className="form-grid quick">
                    {selectedActivities.map((key) => (
                      <Field key={key} label={`Cantidad ${retroActivityLabels[key]}`} type="number" value={row[key]} onChange={(value) => updateRetroRow(index, { [key]: Number(value) } as Partial<RetroRow>)} />
                    ))}
                    <Field label="Hor. inicial" type="number" value={row.horometroInicial} onChange={(value) => updateRetroRow(index, { horometroInicial: Number(value) })} />
                    <Field label="Hor. final" type="number" value={row.horometroFinal} onChange={(value) => updateRetroRow(index, { horometroFinal: Number(value) })} />
                    <Field label="Diesel" type="number" value={row.diesel} onChange={(value) => updateRetroRow(index, { diesel: Number(value) })} />
                  </div>
                  <TextArea label="Observaciones" value={row.observaciones} onChange={(value) => updateRetroRow(index, { observaciones: value })} />
                </article>
              )
            })}
          </>
        )}
      </QuickSection>
      <QuickSection title="Cierre" icon={<ClipboardCheck size={18} />} anchorId="capture-cierre">
        <TextArea label="Comentarios generales" value={record.comentarios} onChange={(value) => setRecord({ ...record, comentarios: value })} />
      </QuickSection>
    </>
  )
}

function SeguridadForm({
  record,
  setRecord,
  showBaseFields = true,
}: {
  record: SeguridadRecord
  setRecord: (record: SeguridadRecord) => void
  showBaseFields?: boolean
}) {
  return (
    <>
      <PanelTitle title="Captura rapida de seguridad" subtitle="Eventos y acciones preventivas" />
      {showBaseFields && (
        <QuickSection title="Turno" icon={<Mountain size={18} />} anchorId="capture-turno">
          <BaseFields record={record} setRecord={setRecord} />
        </QuickSection>
      )}
      <QuickSection title="Resumen" icon={<ShieldCheck size={18} />} anchorId="capture-resumen">
        <div className="form-grid quick">
          <Field label="Accidentes" type="number" value={record.accidentes} onChange={(value) => setRecord({ ...record, accidentes: Number(value) })} />
          <Field label="Incidentes" type="number" value={record.incidentes} onChange={(value) => setRecord({ ...record, incidentes: Number(value) })} />
          <Field label="Fuerza laboral" type="number" value={record.fuerzaLaboral} onChange={(value) => setRecord({ ...record, fuerzaLaboral: Number(value) })} />
        </div>
      </QuickSection>
      <QuickSection title="Hallazgo" icon={<ClipboardCheck size={18} />} anchorId="capture-hallazgo">
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

function QuickSection({ title, icon, children, anchorId }: { title: string; icon: ReactNode; children: ReactNode; anchorId?: string }) {
  return (
    <section className="quick-section" id={anchorId}>
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

async function buildBarrenacionPdf(record: BarrenacionRecord) {
  const pdfDoc = await loadTemplatePdf(BARRENACION_TEMPLATE)
  const page = pdfDoc.getPage(0)
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold)

  drawPdfText(page, font, record.supervisor, 166, 150, 6, 92)
  drawPdfText(page, font, record.fecha, 545, 150, 6, 70)
  drawPdfText(page, bold, 'X', record.turno === '1' ? 688 : 716, 150, 8, 12)

  record.jumbo.filter(hasDrillData).slice(0, 6).forEach((row, index) => {
    drawDrillPdfRow(page, font, row, 184 + index * 12)
  })
  record.maquinaPierna.filter(hasDrillData).slice(0, 6).forEach((row, index) => {
    drawDrillPdfRow(page, font, row, 258 + index * 12)
  })
  record.voladuras.filter(hasBlastData).slice(0, 6).forEach((row, index) => {
    drawBlastPdfRow(page, font, row, 342 + index * 12)
  })

  const activityTop = 420
  const activityRows = [
    record.polvorero,
    record.choferCamion,
    record.choferPipa,
    record.bobCat,
    record.bombeo,
    record.servicios,
  ]
  activityRows.forEach((value, index) => drawPdfText(page, font, value, 167, activityTop + index * 10, 5.8, 190))
  drawPdfWrappedText(page, font, record.comentarios, 478, 416, 125, 5.6, 7, 11)
  drawPdfWrappedText(page, font, record.inasistencias, 610, 416, 120, 5.6, 7, 11)

  return pdfDoc.save()
}

async function buildRezagadoPdf(record: RezagadoRecord) {
  const pdfDoc = await loadTemplatePdf(REZAGADO_TEMPLATE)
  const page = pdfDoc.getPage(0)
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold)

  drawPdfText(page, font, record.supervisor, 140, 76, 6, 120)
  drawPdfText(page, font, record.fecha, 565, 76, 6, 70)
  drawPdfText(page, bold, 'X', record.turno === '1' ? 724 : 748, 76, 8, 12)

  record.scoopTram.filter(hasHaulData).slice(0, 16).forEach((row, index) => {
    drawHaulPdfRow(page, font, row, 114 + index * 10.6)
  })
  record.retro.filter(hasRetroData).slice(0, 14).forEach((row, index) => {
    drawRetroPdfRow(page, font, row, 324 + index * 10.9)
  })

  return pdfDoc.save()
}

function drawDrillPdfRow(page: PdfPage, font: PdfFont, row: DrillRow, top: number) {
  drawPdfText(page, font, row.equipo, 82, top, 5.4, 34)
  drawPdfText(page, font, row.nivelObra, 119, top, 5.4, 48)
  drawPdfText(page, font, row.operador, 169, top, 5.4, 92)
  drawPdfText(page, font, row.ayudante, 270, top, 5.4, 92)
  drawPdfText(page, font, row.rpaCfte, 370, top, 5.2, 19)
  drawPdfText(page, font, valueText(row.longitud), 394, top, 5.2, 17)
  drawPdfText(page, font, valueText(row.cuele), 418, top, 5.2, 17)
  drawPdfText(page, font, valueText(row.desarrollo), 443, top, 5.2, 35)
  drawPdfText(page, font, valueText(row.barrenosDados), 476, top, 5.2, 20)
  drawPdfText(page, font, valueText(row.barrenosCargados), 501, top, 5.2, 20)
  drawPdfText(page, font, valueText(row.metrosDados), 526, top, 5.2, 20)
  drawPdfText(page, font, valueText(row.horasServicio), 550, top, 5.2, 18)
  drawPdfText(page, font, valueText(row.zanco), 566, top, 5.2, 13)
  drawPdfText(page, font, valueText(row.cople), 581, top, 5.2, 13)
  drawPdfText(page, font, valueText(row.barra), 596, top, 5.2, 13)
  drawPdfText(page, font, valueText(row.broca), 611, top, 5.2, 13)
  drawPdfText(page, font, valueText(row.horometroDieselInicial), 632, top, 5.2, 24)
  drawPdfText(page, font, valueText(row.horometroDieselFinal), 661, top, 5.2, 24)
  drawPdfText(page, font, valueText(row.horometroElectInicial), 690, top, 5.2, 24)
  drawPdfText(page, font, valueText(row.horometroElectFinal), 719, top, 5.2, 24)
}

function drawBlastPdfRow(page: PdfPage, font: PdfFont, row: BlastRow, top: number) {
  drawPdfText(page, font, row.obra, 83, top, 5.4, 34)
  drawPdfText(page, font, row.rpaCfte || row.obra, 121, top, 5.4, 48)
  drawPdfText(page, font, row.oficial, 170, top, 5.4, 92)
  drawPdfText(page, font, row.ayudante, 273, top, 5.4, 68)
  drawPdfText(page, font, row.rpaCfte, 349, top, 5.2, 20)
  drawPdfText(page, font, valueText(row.longitud), 376, top, 5.2, 19)
  drawPdfText(page, font, valueText(row.cuele), 398, top, 5.2, 19)
  drawPdfText(page, font, valueText(row.desarrollo), 422, top, 5.2, 35)
  drawPdfText(page, font, valueText(row.barrenosPegados), 447, top, 5.2, 30)
  drawPdfText(page, font, valueText(row.metrosPegados), 488, top, 5.2, 20)
  drawPdfText(page, font, valueText(row.horasServicio), 512, top, 5.2, 20)
  drawPdfText(page, font, valueText(row.anfoInicial), 543, top, 5.2, 40)
  drawPdfText(page, font, valueText(row.anfoFinal), 585, top, 5.2, 40)
  drawPdfText(page, font, row.observaciones, 610, top, 5.2, 125)
}

function drawHaulPdfRow(page: PdfPage, font: PdfFont, row: HaulRow, top: number) {
  drawPdfText(page, font, row.equipo, 29, top, 5.3, 30)
  drawPdfText(page, font, row.operador, 64, top, 5.3, 72)
  drawPdfText(page, font, row.nivelObra, 143, top, 5.3, 68)
  drawPdfText(page, font, row.destino, 218, top, 5.3, 62)
  drawPdfText(page, font, valueText(row.rezagado), 292, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.traspaleo), 320, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.limpia), 348, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.balastreo), 374, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.planilla), 403, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.relleno), 430, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.camiones), 458, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.horometroInicial), 488, top, 5.2, 38)
  drawPdfText(page, font, valueText(row.horometroFinal), 530, top, 5.2, 38)
  drawPdfText(page, font, valueText(row.diesel), 568, top, 5.2, 28)
  drawPdfText(page, font, row.observaciones, 598, top, 5.2, 155)
}

function drawRetroPdfRow(page: PdfPage, font: PdfFont, row: RetroRow, top: number) {
  drawPdfText(page, font, row.equipo, 29, top, 5.3, 30)
  drawPdfText(page, font, row.operador, 64, top, 5.3, 72)
  drawPdfText(page, font, row.nivelObra, 143, top, 5.3, 138)
  drawPdfText(page, font, valueText(row.amacice), 294, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.reAmacice), 321, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.tableo), 349, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.limpia), 375, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.balastreo), 402, top, 5.2, 22)
  drawPdfText(page, font, valueText(row.mtAcequia), 432, top, 5.2, 28)
  drawPdfText(page, font, valueText(row.horometroInicial), 463, top, 5.2, 36)
  drawPdfText(page, font, valueText(row.horometroFinal), 504, top, 5.2, 36)
  drawPdfText(page, font, valueText(row.diesel), 535, top, 5.2, 32)
  drawPdfText(page, font, row.observaciones, 565, top, 5.2, 180)
}

async function buildBarrenacionExcel(record: BarrenacionRecord) {
  const workbook = createWorkbook()
  const sheet = workbook.addWorksheet('Barrenacion y Voladuras')
  setupSheet(sheet, [6, 18, 18, 26, 26, 12, 10, 10, 18, 12, 13, 12, 12, 9, 9, 9, 9, 14, 14, 14, 14])
  await addLogo(workbook, sheet)
  sheet.mergeCells('C2:P2')
  sheet.getCell('C2').value = 'Reporte Diario de Barrenacion y Voladuras'
  sheet.getCell('C2').font = { bold: true, size: 14 }
  sheet.getCell('C2').alignment = centerAlign
  sheet.mergeCells('B4:D4')
  sheet.getCell('B4').value = `Supervisor: ${record.supervisor || ''}`
  sheet.mergeCells('F4:H4')
  sheet.getCell('F4').value = `Fecha: ${record.fecha}`
  sheet.mergeCells('J4:K4')
  sheet.getCell('J4').value = `Turno: ${record.turno}`
  sheet.mergeCells('M4:P4')
  sheet.getCell('M4').value = record.unidad

  const drillHeaders = [
    'Area',
    'Equipo',
    'Nivel / Obra',
    'Operador',
    'Ayudante',
    'RPA/Cfte',
    'Longitud',
    'Cuele',
    'Desb/desc/corte',
    'Barrenos Dados',
    'Barrenos Cargados',
    'Metros Dados',
    'Hrs Servicio',
    'Zanco',
    'Cople',
    'Barra',
    'Broca',
    'Hor. Diesel Inicial',
    'Hor. Diesel Final',
    'Hor. Elect Inicial',
    'Hor. Elect Final',
  ]
  addHeaderRow(sheet, 6, drillHeaders)
  addDrillExcelRows(sheet, 7, 'JUMBO', record.jumbo.filter(hasDrillData), 6)
  addDrillExcelRows(sheet, 13, 'MAQUINA PIERNA', record.maquinaPierna.filter(hasDrillData), 6)

  const blastStart = 21
  const blastHeaders = [
    'Area',
    'Obra',
    'Nivel / Obra',
    'Oficial Voladuras',
    'Ayudante Voladuras',
    'RPA/Cfte',
    'Longitud',
    'Cuele',
    'Desarrollo',
    'Barrenos Pegados',
    'Metros Pegados',
    'Hrs Servicio',
    'ANFO Inicial',
    'ANFO Final',
    'Observaciones',
  ]
  addHeaderRow(sheet, blastStart, blastHeaders)
  addBlastExcelRows(sheet, blastStart + 1, record.voladuras.filter(hasBlastData), 6)

  const activitiesStart = blastStart + 9
  sheet.mergeCells(`A${activitiesStart}:C${activitiesStart}`)
  sheet.getCell(`A${activitiesStart}`).value = 'ACTIVIDADES'
  sheet.mergeCells(`D${activitiesStart}:G${activitiesStart}`)
  sheet.getCell(`D${activitiesStart}`).value = 'RESPONSABLE / COMENTARIO'
  sheet.mergeCells(`H${activitiesStart}:L${activitiesStart}`)
  sheet.getCell(`H${activitiesStart}`).value = 'CONTRATIEMPOS / COMENTARIOS GENERALES'
  sheet.mergeCells(`M${activitiesStart}:Q${activitiesStart}`)
  sheet.getCell(`M${activitiesStart}`).value = 'INASISTENCIAS / PERMISOS'
  const activities = [
    ['POLVORERO', record.polvorero],
    ['CHOFER CAMION PERSONAL', record.choferCamion],
    ['CHOFER PIPA', record.choferPipa],
    ['BOB CAT', record.bobCat],
    ['BOMBEO', record.bombeo],
    ['SERVICIOS', record.servicios],
  ]
  activities.forEach(([label, value], index) => {
    const rowNumber = activitiesStart + 1 + index
    sheet.mergeCells(`A${rowNumber}:C${rowNumber}`)
    sheet.getCell(`A${rowNumber}`).value = label
    sheet.mergeCells(`D${rowNumber}:G${rowNumber}`)
    sheet.getCell(`D${rowNumber}`).value = value
  })
  sheet.mergeCells(`H${activitiesStart + 1}:L${activitiesStart + 6}`)
  sheet.getCell(`H${activitiesStart + 1}`).value = record.comentarios
  sheet.mergeCells(`M${activitiesStart + 1}:Q${activitiesStart + 6}`)
  sheet.getCell(`M${activitiesStart + 1}`).value = record.inasistencias
  styleUsedCells(sheet, 1, activitiesStart + 6, 21)
  return workbook.xlsx.writeBuffer()
}

async function buildRezagadoExcel(record: RezagadoRecord) {
  const workbook = createWorkbook()
  const sheet = workbook.addWorksheet('Rezagado')
  setupSheet(sheet, [8, 18, 22, 22, 18, 11, 11, 11, 12, 11, 11, 11, 14, 14, 12, 36])
  await addLogo(workbook, sheet)
  sheet.mergeCells('C2:N2')
  sheet.getCell('C2').value = 'Reporte Diario de Rezagado'
  sheet.getCell('C2').font = { bold: true, size: 14 }
  sheet.getCell('C2').alignment = centerAlign
  sheet.mergeCells('B4:D4')
  sheet.getCell('B4').value = `Supervisor: ${record.supervisor || ''}`
  sheet.mergeCells('F4:H4')
  sheet.getCell('F4').value = `Fecha: ${record.fecha}`
  sheet.mergeCells('J4:K4')
  sheet.getCell('J4').value = `Turno: ${record.turno}`

  const haulHeaders = [
    'Area',
    'Equipo',
    'Operador',
    'Nivel / Obra',
    'Destino',
    'Rezagado',
    'Traspaleo',
    'Limpia',
    'Balastreo',
    'Planilla',
    'Relleno',
    'Camiones',
    'Hor. Inicial',
    'Hor. Final',
    'Diesel',
    'Observaciones',
  ]
  addHeaderRow(sheet, 6, haulHeaders)
  addHaulExcelRows(sheet, 7, record.scoopTram.filter(hasHaulData), 16)

  const retroStart = 25
  const retroHeaders = [
    'Area',
    'Equipo',
    'Operador',
    'Nivel / Obra',
    'Amacice',
    'Re-amacice',
    'Tableo',
    'Limpia',
    'Balastreo',
    'MT Acequia',
    'Hor. Inicial',
    'Hor. Final',
    'Diesel',
    'Observaciones',
  ]
  addHeaderRow(sheet, retroStart, retroHeaders)
  addRetroExcelRows(sheet, retroStart + 1, record.retro.filter(hasRetroData), 14)
  styleUsedCells(sheet, 1, retroStart + 14, 16)
  return workbook.xlsx.writeBuffer()
}

function addDrillExcelRows(sheet: ExcelJS.Worksheet, startRow: number, area: string, rows: DrillRow[], minimumRows: number) {
  const totalRows = Math.max(rows.length, minimumRows)
  if (totalRows > 1) sheet.mergeCells(startRow, 1, startRow + totalRows - 1, 1)
  sheet.getCell(startRow, 1).value = area
  for (let index = 0; index < totalRows; index += 1) {
    const row = rows[index]
    const values = row
      ? [
          row.equipo,
          row.nivelObra,
          row.operador,
          row.ayudante,
          row.rpaCfte,
          row.longitud,
          row.cuele,
          row.desarrollo,
          row.barrenosDados,
          row.barrenosCargados,
          row.metrosDados,
          row.horasServicio,
          row.zanco,
          row.cople,
          row.barra,
          row.broca,
          row.horometroDieselInicial,
          row.horometroDieselFinal,
          row.horometroElectInicial,
          row.horometroElectFinal,
        ]
      : []
    values.forEach((value, colIndex) => {
      sheet.getCell(startRow + index, colIndex + 2).value = value || ''
    })
  }
}

function addBlastExcelRows(sheet: ExcelJS.Worksheet, startRow: number, rows: BlastRow[], minimumRows: number) {
  const totalRows = Math.max(rows.length, minimumRows)
  if (totalRows > 1) sheet.mergeCells(startRow, 1, startRow + totalRows - 1, 1)
  sheet.getCell(startRow, 1).value = 'VOLADURAS'
  for (let index = 0; index < totalRows; index += 1) {
    const row = rows[index]
    const values = row
      ? [
          row.obra,
          row.rpaCfte || row.obra,
          row.oficial,
          row.ayudante,
          row.rpaCfte,
          row.longitud,
          row.cuele,
          row.desarrollo,
          row.barrenosPegados,
          row.metrosPegados,
          row.horasServicio,
          row.anfoInicial,
          row.anfoFinal,
          row.observaciones,
        ]
      : []
    values.forEach((value, colIndex) => {
      sheet.getCell(startRow + index, colIndex + 2).value = value || ''
    })
  }
}

function addHaulExcelRows(sheet: ExcelJS.Worksheet, startRow: number, rows: HaulRow[], minimumRows: number) {
  const totalRows = Math.max(rows.length, minimumRows)
  if (totalRows > 1) sheet.mergeCells(startRow, 1, startRow + totalRows - 1, 1)
  sheet.getCell(startRow, 1).value = 'SCOOP TRAM'
  for (let index = 0; index < totalRows; index += 1) {
    const row = rows[index]
    const values = row
      ? [
          row.equipo,
          row.operador,
          row.nivelObra,
          row.destino,
          row.rezagado,
          row.traspaleo,
          row.limpia,
          row.balastreo,
          row.planilla,
          row.relleno,
          row.camiones,
          row.horometroInicial,
          row.horometroFinal,
          row.diesel,
          row.observaciones,
        ]
      : []
    values.forEach((value, colIndex) => {
      sheet.getCell(startRow + index, colIndex + 2).value = value || ''
    })
  }
}

function addRetroExcelRows(sheet: ExcelJS.Worksheet, startRow: number, rows: RetroRow[], minimumRows: number) {
  const totalRows = Math.max(rows.length, minimumRows)
  if (totalRows > 1) sheet.mergeCells(startRow, 1, startRow + totalRows - 1, 1)
  sheet.getCell(startRow, 1).value = 'RETRO'
  for (let index = 0; index < totalRows; index += 1) {
    const row = rows[index]
    const values = row
      ? [
          row.equipo,
          row.operador,
          row.nivelObra,
          row.amacice,
          row.reAmacice,
          row.tableo,
          row.limpia,
          row.balastreo,
          row.mtAcequia,
          row.horometroInicial,
          row.horometroFinal,
          row.diesel,
          row.observaciones,
        ]
      : []
    values.forEach((value, colIndex) => {
      sheet.getCell(startRow + index, colIndex + 2).value = value || ''
    })
  }
}

function loadRecords(): MineRecord[] {
  try {
    return (JSON.parse(localStorage.getItem(STORAGE_KEY) ?? '[]') as MineRecord[]).map(normalizeRecord)
  } catch {
    return []
  }
}

function loadCaptureDraft(): CaptureDraft | null {
  try {
    const stored = localStorage.getItem(CAPTURE_DRAFT_KEY)
    if (!stored) return null
    const draft = JSON.parse(stored) as CaptureDraft
    const normalized: CaptureDraft = {
      captureMode: isCaptureMode(draft.captureMode) ? draft.captureMode : 'modulo',
      captureLayout: isCaptureLayout(draft.captureLayout) ? draft.captureLayout : 'asistente',
      formType: isRecordType(draft.formType) ? draft.formType : 'barrenacion',
      turnoBase: normalizeShiftBase(draft.turnoBase),
      turnoModules: normalizeTurnoModules(draft.turnoModules),
      barrenacion: normalizeRecord(draft.barrenacion ?? makeBarrenacion()) as BarrenacionRecord,
      rezagado: normalizeRecord(draft.rezagado ?? makeRezagado()) as RezagadoRecord,
      seguridad: normalizeRecord(draft.seguridad ?? makeSeguridad()) as SeguridadRecord,
      editingId: draft.editingId ?? null,
      savedAt: draft.savedAt ?? nowIso(),
    }
    return hasCaptureDraftContent(normalized) ? normalized : null
  } catch {
    return null
  }
}

function loadChatMessages(): ChatMessage[] {
  try {
    return mergeChatMessages(
      (JSON.parse(localStorage.getItem(MESSAGE_STORAGE_KEY) ?? '[]') as ChatMessage[]).map(normalizeChatMessage),
    )
  } catch {
    return []
  }
}

function getInitialSection(): AppSection {
  const params = new URLSearchParams(window.location.search)
  const requestedView = params.get('vista')
  const isApk = Capacitor.isNativePlatform()
  if (requestedView === 'inicio' || requestedView === 'home') return 'home'
  if (requestedView === 'captura') return 'captura'
  if (requestedView === 'historial' || requestedView === 'bitacora') return 'historial'
  if (isApk && (requestedView === 'revision' || requestedView === 'dashboard')) return 'historial'
  if (!isApk && (requestedView === 'revision' || requestedView === 'dashboard')) return 'dashboard'
  return isApk ? 'home' : 'dashboard'
}

function loadApiUrl() {
  return localStorage.getItem(API_URL_KEY) ?? import.meta.env.VITE_API_URL ?? DEFAULT_API_URL
}

function loadWhatsAppNumber() {
  return localStorage.getItem(WHATSAPP_NUMBER_KEY) ?? ''
}

function loadOperatorMode() {
  return localStorage.getItem(OPERATOR_MODE_KEY) === '1'
}

function normalizeWhatsAppNumber(value: string) {
  const digits = value.replace(/\D/g, '')
  if (digits.length === 10) return `52${digits}`
  if (digits.length >= 11 && digits.length <= 15) return digits
  return ''
}

function resolveApiBase(apiUrl: string) {
  return (apiUrl.trim() || DEFAULT_API_URL).replace(/\/$/, '')
}

function normalizeShiftBase(base?: Partial<ShiftBase>): ShiftBase {
  return {
    supervisor: base?.supervisor ?? '',
    fecha: base?.fecha ?? today,
    turno: base?.turno === '2' ? '2' : '1',
    unidad: base?.unidad ?? baseDefaults.unidad,
  }
}

function normalizeTurnoModules(modules?: Partial<TurnoModules>): TurnoModules {
  return {
    barrenacion: modules?.barrenacion ?? defaultTurnoModules.barrenacion,
    rezagado: modules?.rezagado ?? defaultTurnoModules.rezagado,
    seguridad: modules?.seguridad ?? defaultTurnoModules.seguridad,
  }
}

function applyShiftBase<T extends MineRecord>(record: T, base: ShiftBase, updatedAt: string): T {
  return normalizeRecord({
    ...record,
    ...base,
    updatedAt,
    deletedAt: undefined,
    syncedAt: undefined,
  } as MineRecord) as T
}

function normalizeRecord(record: MineRecord): MineRecord {
  const createdAt = record.createdAt ?? nowIso()
  const baseRecord = {
    ...record,
    createdAt,
    updatedAt: record.updatedAt ?? createdAt,
  } as MineRecord

  if (baseRecord.type === 'barrenacion') {
    return {
      ...baseRecord,
      activeActivity: getBarrenacionActivity(baseRecord),
    }
  }
  if (baseRecord.type === 'rezagado') {
    const scoopRow = baseRecord.scoopTram[0] ?? emptyHaulRow(defaultScoop)
    const retroRow = baseRecord.retro[0] ?? emptyRetroRow()
    return {
      ...baseRecord,
      activeEquipment: getRezagadoEquipment(baseRecord),
      activeHaulActivity: getHaulActivity(baseRecord, scoopRow),
      activeRetroActivity: getRetroActivity(baseRecord, retroRow),
    }
  }
  return baseRecord
}

function normalizeChatMessage(message: ChatMessage): ChatMessage {
  const createdAt = message.createdAt ?? nowIso()
  return {
    ...message,
    id: message.id || crypto.randomUUID(),
    type: ['aviso', 'mensaje', 'urgente'].includes(message.type) ? message.type : 'mensaje',
    author: message.author?.trim() || 'Operacion',
    text: message.text?.trim() || '',
    createdAt,
    updatedAt: message.updatedAt ?? createdAt,
  }
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

function mergeChatMessages(messages: ChatMessage[]) {
  const map = new Map<string, ChatMessage>()
  messages.map(normalizeChatMessage).filter((item) => item.text).forEach((item) => {
    const current = map.get(item.id)
    if (!current || new Date(item.updatedAt).getTime() >= new Date(current.updatedAt).getTime()) {
      map.set(item.id, item)
    }
  })
  return Array.from(map.values()).sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())
}

function buildEquipmentFavorites(records: MineRecord[]): EquipmentFavorites {
  const counters = {
    jumbo: new Map<string, number>(),
    scoop: new Map<string, number>(),
    retro: new Map<string, number>(),
  }

  function count(target: keyof EquipmentFavorites, equipo: string) {
    const clean = equipo.trim()
    if (!clean) return
    counters[target].set(clean, (counters[target].get(clean) ?? 0) + 1)
  }

  records.forEach((record) => {
    if (record.type === 'barrenacion') {
      record.jumbo.forEach((row) => count('jumbo', row.equipo))
      record.maquinaPierna.forEach((row) => count('jumbo', row.equipo))
    }
    if (record.type === 'rezagado') {
      record.scoopTram.forEach((row) => count('scoop', row.equipo))
      record.retro.forEach((row) => count('retro', row.equipo))
    }
  })

  const sorted = (map: Map<string, number>) =>
    Array.from(map.entries())
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .slice(0, 8)
      .map(([value]) => value)

  return {
    jumbo: sorted(counters.jumbo),
    scoop: sorted(counters.scoop),
    retro: sorted(counters.retro),
  }
}

function prioritizeOptions(options: string[], selected: string, favorites: string[]) {
  const seen = new Set<string>()
  const next: string[] = []
  ;[selected, ...favorites, ...options].forEach((item) => {
    const value = item.trim()
    if (!value || seen.has(value)) return
    seen.add(value)
    next.push(value)
  })
  return next
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

function buildCaptureReview(
  captureMode: CaptureMode,
  formType: RecordType,
  turnoBase: ShiftBase,
  turnoModules: TurnoModules,
  barrenacion: BarrenacionRecord,
  rezagado: RezagadoRecord,
  seguridad: SeguridadRecord,
  editingId: string | null,
): CaptureReview {
  const records = getRecordsForReview(captureMode, formType, turnoBase, turnoModules, barrenacion, rezagado, seguridad, editingId)
  const warnings: string[] = []
  const errors: string[] = []
  if (records.length === 0) errors.push('Selecciona al menos un modulo antes de guardar.')
  records.forEach((record) => validateReviewRecord(record, warnings))
  return {
    title: captureMode === 'turno' && !editingId ? 'Turno completo' : editingId ? 'Guardar cambios' : 'Nueva captura',
    records,
    warnings,
    errors,
    kpis: computeKpis(records),
  }
}

function getRecordsForReview(
  captureMode: CaptureMode,
  formType: RecordType,
  turnoBase: ShiftBase,
  turnoModules: TurnoModules,
  barrenacion: BarrenacionRecord,
  rezagado: RezagadoRecord,
  seguridad: SeguridadRecord,
  editingId: string | null,
) {
  const updatedAt = nowIso()
  if (captureMode === 'turno' && !editingId) {
    const selected: MineRecord[] = []
    if (turnoModules.barrenacion) selected.push(applyShiftBase(barrenacion, turnoBase, updatedAt))
    if (turnoModules.rezagado) selected.push(applyShiftBase(rezagado, turnoBase, updatedAt))
    if (turnoModules.seguridad) selected.push(applyShiftBase(seguridad, turnoBase, updatedAt))
    return selected
  }
  const source = formType === 'barrenacion' ? barrenacion : formType === 'rezagado' ? rezagado : seguridad
  return [normalizeRecord({ ...source, updatedAt, deletedAt: undefined, syncedAt: undefined } as MineRecord)]
}

function validateReviewRecord(record: MineRecord, warnings: string[]) {
  const label = labelType(record.type)
  if (!record.supervisor.trim()) warnings.push(`${label}: falta supervisor.`)
  if (!record.fecha) warnings.push(`${label}: falta fecha.`)
  if (record.type === 'barrenacion') validateBarrenacionReview(record, warnings)
  if (record.type === 'rezagado') validateRezagadoReview(record, warnings)
  if (record.type === 'seguridad') validateSeguridadReview(record, warnings)
}

function validateBarrenacionReview(record: BarrenacionRecord, warnings: string[]) {
  const drillRows = [...record.jumbo, ...record.maquinaPierna].filter(hasDrillData)
  const blastRows = record.voladuras.filter(hasBlastData)
  if (drillRows.length === 0 && blastRows.length === 0) {
    warnings.push('Barrenacion: no hay equipos, frentes o voladuras con produccion capturada.')
  }
  drillRows.forEach((row, index) => {
    if (!row.operador.trim()) warnings.push(`Barrenacion equipo ${index + 1}: falta operador.`)
    if (!row.nivelObra.trim()) warnings.push(`Barrenacion equipo ${index + 1}: falta nivel/obra.`)
    if (row.horometroDieselFinal > 0 && row.horometroDieselFinal < row.horometroDieselInicial) warnings.push(`Barrenacion equipo ${index + 1}: horometro diesel final menor al inicial.`)
    if (row.horometroElectFinal > 0 && row.horometroElectFinal < row.horometroElectInicial) warnings.push(`Barrenacion equipo ${index + 1}: horometro electrico final menor al inicial.`)
    if (!row.barrenosDados && !row.barrenosCargados && !row.metrosDados) warnings.push(`Barrenacion equipo ${index + 1}: produccion en cero.`)
  })
  blastRows.forEach((row, index) => {
    if (!row.obra.trim()) warnings.push(`Voladura ${index + 1}: falta obra/frente.`)
    if (!row.oficial.trim()) warnings.push(`Voladura ${index + 1}: falta oficial.`)
    if (!row.barrenosPegados && !row.metrosPegados) warnings.push(`Voladura ${index + 1}: produccion en cero.`)
  })
}

function validateRezagadoReview(record: RezagadoRecord, warnings: string[]) {
  const scoopRows = record.scoopTram.filter(hasHaulData)
  const retroRows = record.retro.filter(hasRetroData)
  if (scoopRows.length === 0 && retroRows.length === 0) warnings.push('Rezagado: no hay equipos con produccion capturada.')
  scoopRows.forEach((row, index) => {
    if (!row.operador.trim()) warnings.push(`Scoop ${index + 1}: falta operador.`)
    if (!row.nivelObra.trim()) warnings.push(`Scoop ${index + 1}: falta nivel/obra.`)
    if (row.horometroFinal > 0 && row.horometroFinal < row.horometroInicial) warnings.push(`Scoop ${index + 1}: horometro final menor al inicial.`)
    if (!row.rezagado && !row.traspaleo && !row.limpia && !row.balastreo && !row.planilla && !row.relleno) warnings.push(`Scoop ${index + 1}: trabajo en cero.`)
  })
  retroRows.forEach((row, index) => {
    if (!row.operador.trim()) warnings.push(`Retro ${index + 1}: falta operador.`)
    if (!row.nivelObra.trim()) warnings.push(`Retro ${index + 1}: falta nivel/obra.`)
    if (row.horometroFinal > 0 && row.horometroFinal < row.horometroInicial) warnings.push(`Retro ${index + 1}: horometro final menor al inicial.`)
    if (!row.amacice && !row.reAmacice && !row.tableo && !row.limpia && !row.balastreo && !row.mtAcequia) warnings.push(`Retro ${index + 1}: trabajo en cero.`)
  })
}

function validateSeguridadReview(record: SeguridadRecord, warnings: string[]) {
  const hasText = Boolean(record.actosInseguros || record.condicionesInseguras || record.platicaSeguridad || record.actividadesSeguridad || record.correccionesMejoras || record.actividadesOperacion || record.observaciones)
  if (!record.accidentes && !record.incidentes && !record.fuerzaLaboral && !hasText) {
    warnings.push('Seguridad: no hay resumen, hallazgo o fuerza laboral capturada.')
  }
}

function labelType(type: RecordType) {
  return {
    barrenacion: 'Barrenacion y voladuras',
    rezagado: 'Rezagado retro',
    seguridad: 'Seguridad',
  }[type]
}

function shortTypeLabel(type: RecordType) {
  return {
    barrenacion: 'Barrenacion',
    rezagado: 'Rezagado',
    seguridad: 'Seguridad',
  }[type]
}

function shortBarrenacionActivityLabel(type: BarrenacionActivity) {
  return {
    jumbo: 'Jumbo',
    maquinaPierna: 'Maquina pierna',
    voladura: 'Voladura',
  }[type]
}

function labelChatType(type: ChatMessageType) {
  return {
    aviso: 'Aviso',
    mensaje: 'Mensaje',
    urgente: 'Urgente',
  }[type]
}

function formatDateTime(value: string) {
  return new Date(value).toLocaleString('es-MX', {
    dateStyle: 'short',
    timeStyle: 'short',
  })
}

async function loadTemplatePdf(path: string) {
  const response = await fetch(path)
  if (!response.ok) throw new Error(`No se pudo cargar la plantilla ${path}`)
  return PDFDocument.load(await response.arrayBuffer())
}

function drawPdfText(
  page: PdfPage,
  font: PdfFont,
  value: string | number | undefined,
  x: number,
  top: number,
  size: number,
  maxWidth: number,
) {
  const text = fitPdfText(font, valueText(value, true), size, maxWidth)
  if (!text) return
  page.drawText(text, {
    x,
    y: page.getHeight() - top - size,
    size,
    font,
    color: rgb(0, 0, 0),
  })
}

function drawPdfWrappedText(
  page: PdfPage,
  font: PdfFont,
  value: string | number | undefined,
  x: number,
  top: number,
  maxWidth: number,
  size: number,
  lineHeight: number,
  maxLines: number,
) {
  const text = valueText(value, true)
  if (!text) return
  const lines = wrapPdfText(font, text, size, maxWidth).slice(0, maxLines)
  lines.forEach((line, index) => {
    drawPdfText(page, font, line, x, top + index * lineHeight, size, maxWidth)
  })
}

function wrapPdfText(font: PdfFont, text: string, size: number, maxWidth: number) {
  const words = text.split(/\s+/).filter(Boolean)
  const lines: string[] = []
  let current = ''
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word
    if (font.widthOfTextAtSize(next, size) <= maxWidth) {
      current = next
      return
    }
    if (current) lines.push(current)
    current = word
  })
  if (current) lines.push(current)
  return lines
}

function fitPdfText(font: PdfFont, text: string, size: number, maxWidth: number) {
  if (font.widthOfTextAtSize(text, size) <= maxWidth) return text
  let output = text
  while (output.length > 0 && font.widthOfTextAtSize(`${output}...`, size) > maxWidth) {
    output = output.slice(0, -1)
  }
  return output ? `${output}...` : ''
}

function valueText(value: string | number | undefined, showZero = false) {
  if (value === undefined || value === null) return ''
  if (typeof value === 'number') {
    if (!showZero && value === 0) return ''
    return Number.isInteger(value) ? String(value) : value.toFixed(2)
  }
  return value
}

function downloadBlob(content: BlobPart | Uint8Array, filename: string, type: string) {
  const blob = new Blob([content instanceof Uint8Array ? (content.slice().buffer as ArrayBuffer) : content], { type })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

function createWorkbook() {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'MGA Bitacora Mina'
  workbook.created = new Date()
  workbook.modified = new Date()
  return workbook
}

function setupSheet(sheet: ExcelJS.Worksheet, widths: number[]) {
  sheet.views = [{ showGridLines: false }]
  sheet.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    paperSize: 9,
  }
  sheet.columns = widths.map((width) => ({ width }))
  for (let rowNumber = 1; rowNumber <= 44; rowNumber += 1) {
    sheet.getRow(rowNumber).height = 18
  }
}

async function addLogo(workbook: ExcelJS.Workbook, sheet: ExcelJS.Worksheet) {
  try {
    const response = await fetch('/mga-logo.jfif')
    if (!response.ok) return
    const base64 = await blobToDataUrl(await response.blob())
    const imageId = workbook.addImage({ base64, extension: 'jpeg' })
    sheet.addImage(imageId, {
      tl: { col: 0.1, row: 0.2 },
      ext: { width: 90, height: 44 },
    })
  } catch {
    // Logo is decorative in Excel export; keep the workbook usable if it cannot be embedded.
  }
}

function blobToDataUrl(blob: Blob) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result))
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(blob)
  })
}

function addHeaderRow(sheet: ExcelJS.Worksheet, rowNumber: number, headers: string[]) {
  const row = sheet.getRow(rowNumber)
  row.values = headers
  row.font = { bold: true, size: 9 }
  row.alignment = centerAlign
  row.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } }
  })
}

function styleUsedCells(sheet: ExcelJS.Worksheet, firstRow: number, lastRow: number, lastColumn: number) {
  for (let rowNumber = firstRow; rowNumber <= lastRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber)
    for (let colNumber = 1; colNumber <= lastColumn; colNumber += 1) {
      const cell = row.getCell(colNumber)
      cell.border = thinBorder
      cell.alignment = {
        vertical: 'middle',
        horizontal: colNumber === 1 ? 'center' : 'center',
        wrapText: true,
      }
      if (colNumber === 1) cell.font = { ...cell.font, bold: true }
    }
  }
}

const thinBorder = {
  top: { style: 'thin' as const },
  left: { style: 'thin' as const },
  bottom: { style: 'thin' as const },
  right: { style: 'thin' as const },
}

const centerAlign = {
  vertical: 'middle' as const,
  horizontal: 'center' as const,
  wrapText: true,
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
  if (isBarrenacionActivity(record.activeActivity)) return record.activeActivity
  if (record.voladuras.some(hasBlastData)) return 'voladura'
  if (record.maquinaPierna.some(hasDrillData)) return 'maquinaPierna'
  return 'jumbo'
}

function getRezagadoEquipment(record: RezagadoRecord): RezagadoEquipment {
  if (isRezagadoEquipment(record.activeEquipment)) return record.activeEquipment
  if (record.retro.some(hasRetroData) && !record.scoopTram.some(hasHaulData)) return 'retro'
  return 'scoop'
}

function getHaulActivity(record: RezagadoRecord, row: HaulRow): HaulActivity {
  if (isHaulActivity(record.activeHaulActivity)) return record.activeHaulActivity
  return getActiveKey(row, haulActivityKeys, 'rezagado')
}

function getRetroActivity(record: RezagadoRecord, row: RetroRow): RetroActivity {
  if (isRetroActivity(record.activeRetroActivity)) return record.activeRetroActivity
  return getActiveKey(row, retroActivityKeys, 'amacice')
}

function getSelectedHaulActivities(row: HaulRow): HaulActivity[] {
  const stored = (row.selectedActivities ?? []).filter(isHaulActivity)
  const withValues = haulActivityKeys.filter((key) => Number(row[key] || 0) > 0)
  const selected = Array.from(new Set<HaulActivity>([...stored, ...withValues]))
  return selected.length ? selected : ['rezagado']
}

function getSelectedRetroActivities(row: RetroRow): RetroActivity[] {
  const stored = (row.selectedActivities ?? []).filter(isRetroActivity)
  const withValues = retroActivityKeys.filter((key) => Number(row[key] || 0) > 0)
  const selected = Array.from(new Set<RetroActivity>([...stored, ...withValues]))
  return selected.length ? selected : ['amacice']
}

function hasCaptureDraftContent(draft: CaptureDraft) {
  return Boolean(
    draft.editingId
      || hasShiftBaseData(draft.turnoBase)
      || modulesChanged(draft.turnoModules)
      || hasRecordDraftData(draft.barrenacion)
      || hasRecordDraftData(draft.rezagado)
      || hasRecordDraftData(draft.seguridad),
  )
}

function hasShiftBaseData(base: ShiftBase) {
  return Boolean(
    base.supervisor.trim()
      || base.fecha !== today
      || base.turno !== baseDefaults.turno
      || base.unidad.trim() !== baseDefaults.unidad,
  )
}

function modulesChanged(modules: TurnoModules) {
  return modules.barrenacion !== defaultTurnoModules.barrenacion
    || modules.rezagado !== defaultTurnoModules.rezagado
    || modules.seguridad !== defaultTurnoModules.seguridad
}

function hasRecordDraftData(record: MineRecord) {
  if (hasShiftBaseData(record)) return true
  if (record.type === 'barrenacion') {
    return Boolean(
      record.jumbo.some(hasDrillData)
        || record.maquinaPierna.some(hasDrillData)
        || record.voladuras.some(hasBlastData)
        || record.polvorero
        || record.choferCamion
        || record.choferPipa
        || record.bobCat
        || record.bombeo
        || record.servicios
        || record.comentarios
        || record.inasistencias,
    )
  }
  if (record.type === 'rezagado') {
    return Boolean(record.scoopTram.some(hasHaulData) || record.retro.some(hasRetroData) || record.comentarios)
  }
  return Boolean(
    record.accidentes
      || record.incidentes
      || record.fuerzaLaboral
      || record.actosInseguros
      || record.condicionesInseguras
      || record.platicaSeguridad
      || record.actividadesSeguridad
      || record.correccionesMejoras
      || record.actividadesOperacion
      || record.observaciones,
  )
}

function isCaptureMode(value: unknown): value is CaptureMode {
  return value === 'modulo' || value === 'turno'
}

function isCaptureLayout(value: unknown): value is CaptureLayout {
  return value === 'asistente' || value === 'formulario'
}

function isRecordType(value: unknown): value is RecordType {
  return value === 'barrenacion' || value === 'rezagado' || value === 'seguridad'
}

function isBarrenacionActivity(value: unknown): value is BarrenacionActivity {
  return barrenacionActivities.includes(value as BarrenacionActivity)
}

function isRezagadoEquipment(value: unknown): value is RezagadoEquipment {
  return rezagadoEquipments.includes(value as RezagadoEquipment)
}

function isHaulActivity(value: unknown): value is HaulActivity {
  return haulActivityKeys.includes(value as HaulActivity)
}

function isRetroActivity(value: unknown): value is RetroActivity {
  return retroActivityKeys.includes(value as RetroActivity)
}

function getActiveKey<T, K extends keyof T>(
  row: T,
  keys: readonly K[],
  fallback: K,
) {
  return keys.find((key) => Number(row[key] || 0) > 0) ?? fallback
}

export default App
