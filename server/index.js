import cors from 'cors'
import express from 'express'
import { randomUUID } from 'node:crypto'
import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const app = express()
const port = process.env.PORT || 3000
const dataFile =
  process.env.DATA_FILE || path.join(__dirname, '..', 'data', 'records.json')
const distDir = path.join(__dirname, '..', 'dist')
const defaultPublicUrl = 'https://mga-bitacora-mina.onrender.com'

app.use(cors())
app.use(express.json({ limit: '15mb' }))

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, service: 'mga-bitacora-mina' })
})

app.get('/api/records', async (_req, res) => {
  res.json({ records: await readRecords() })
})

app.post('/api/records/sync', async (req, res) => {
  const incoming = Array.isArray(req.body?.records) ? req.body.records : []
  const existing = await readRecords()
  const slackEvents = getSlackEvents(incoming, existing)
  const records = mergeRecords([...existing, ...incoming])
  await writeRecords(records)
  void notifySlack(slackEvents)
  res.json({ records })
})

app.put('/api/records/:id', async (req, res) => {
  const now = new Date().toISOString()
  const existing = await readRecords()
  const record = normalizeRecord({ ...req.body, id: req.params.id, updatedAt: now })
  const slackEvents = getSlackEvents([record], existing)
  const records = mergeRecords([...existing, record])
  await writeRecords(records)
  void notifySlack(slackEvents)
  res.json({ record })
})

app.delete('/api/records/:id', async (req, res) => {
  const now = new Date().toISOString()
  const existing = await readRecords()
  const current = existing.find((record) => record.id === req.params.id)
  const deleted = normalizeRecord({
    ...(current ?? { id: req.params.id, type: 'seguridad', createdAt: now }),
    deletedAt: now,
    updatedAt: now,
  })
  const records = mergeRecords([...existing, deleted])
  await writeRecords(records)
  void notifySlack([{ action: 'deleted', record: deleted }])
  res.json({ record: deleted })
})

app.use(express.static(distDir))
app.get(/.*/, (_req, res) => {
  res.sendFile(path.join(distDir, 'index.html'))
})

app.listen(port, '0.0.0.0', () => {
  console.log(`MGA bitacora server running on ${port}`)
})

async function readRecords() {
  try {
    const content = await fs.readFile(dataFile, 'utf8')
    return mergeRecords(JSON.parse(content))
  } catch (error) {
    if (error.code !== 'ENOENT') console.error(error)
    return []
  }
}

async function writeRecords(records) {
  await fs.mkdir(path.dirname(dataFile), { recursive: true })
  await fs.writeFile(dataFile, JSON.stringify(mergeRecords(records), null, 2))
}

function normalizeRecord(record) {
  const now = new Date().toISOString()
  const createdAt = record.createdAt || now
  return {
    ...record,
    id: record.id || randomUUID(),
    createdAt,
    updatedAt: record.updatedAt || createdAt,
  }
}

function mergeRecords(records) {
  const map = new Map()
  records.map(normalizeRecord).forEach((record) => {
    const current = map.get(record.id)
    if (!current || new Date(record.updatedAt).getTime() >= new Date(current.updatedAt).getTime()) {
      map.set(record.id, record)
    }
  })
  return Array.from(map.values()).sort(
    (a, b) => new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime(),
  )
}

function getSlackEvents(incoming, existing) {
  const existingById = new Map(existing.map((record) => [record.id, record]))
  return incoming
    .map(normalizeRecord)
    .filter((record) => {
      const current = existingById.get(record.id)
      return !current || new Date(record.updatedAt).getTime() > new Date(current.updatedAt).getTime()
    })
    .map((record) => ({
      action: record.deletedAt ? 'deleted' : existingById.has(record.id) ? 'updated' : 'created',
      record,
    }))
}

async function notifySlack(events) {
  const webhookUrl = process.env.SLACK_WEBHOOK_URL
  if (!webhookUrl || events.length === 0) return

  const maxEvents = 6
  const lines = events.slice(0, maxEvents).map(formatSlackEvent)
  const extra = events.length > maxEvents ? `\n...y ${events.length - maxEvents} cambio(s) mas.` : ''
  const publicUrl = process.env.PUBLIC_APP_URL || defaultPublicUrl
  const text = [
    `MGA Bitacora Mina: ${events.length} cambio(s) sincronizado(s)`,
    ...lines,
    extra,
    `Panel: ${publicUrl}`,
  ]
    .filter(Boolean)
    .join('\n')

  try {
    const response = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text }),
    })
    if (!response.ok) {
      console.error(`Slack webhook failed with HTTP ${response.status}`)
    }
  } catch (error) {
    console.error('Slack webhook failed', error)
  }
}

function formatSlackEvent({ action, record }) {
  const actionLabel = {
    created: 'Nuevo',
    updated: 'Actualizado',
    deleted: 'Eliminado',
  }[action]
  const typeLabel = {
    barrenacion: 'Barrenacion/voladuras',
    rezagado: 'Rezagado',
    seguridad: 'Seguridad',
  }[record.type] || record.type
  const details = [
    record.fecha,
    record.turno ? `Turno ${record.turno}` : '',
    record.supervisor ? `Supervisor: ${record.supervisor}` : '',
    record.unidad,
    getRecordSummary(record),
  ].filter(Boolean)
  return `- ${actionLabel}: ${typeLabel} | ${details.join(' | ')}`
}

function getRecordSummary(record) {
  if (record.type === 'barrenacion') {
    const jumboCount = Array.isArray(record.jumbo) ? record.jumbo.length : 0
    const piernaCount = Array.isArray(record.maquinaPierna) ? record.maquinaPierna.length : 0
    const voladuraCount = Array.isArray(record.voladuras) ? record.voladuras.length : 0
    return `Jumbo: ${jumboCount}, Maq. pierna: ${piernaCount}, Voladuras: ${voladuraCount}`
  }
  if (record.type === 'rezagado') {
    const scoopCount = Array.isArray(record.scoopTram) ? record.scoopTram.length : 0
    const retroCount = Array.isArray(record.retro) ? record.retro.length : 0
    return `Scoop: ${scoopCount}, Retro: ${retroCount}`
  }
  if (record.type === 'seguridad') {
    const accidentes = Number(record.accidentes || 0)
    const incidentes = Number(record.incidentes || 0)
    const fuerzaLaboral = Number(record.fuerzaLaboral || 0)
    return `Accidentes: ${accidentes}, Incidentes: ${incidentes}, Fuerza laboral: ${fuerzaLaboral}`
  }
  return ''
}
