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
const messagesFile =
  process.env.MESSAGES_FILE || path.join(path.dirname(dataFile), 'messages.json')
const usersFile =
  process.env.USERS_FILE || path.join(path.dirname(dataFile), 'users.json')
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

app.get('/api/messages', async (_req, res) => {
  res.json({ messages: await readMessages() })
})

app.post('/api/messages/sync', async (req, res) => {
  const incoming = Array.isArray(req.body?.messages) ? req.body.messages : []
  const existing = await readMessages()
  const slackEvents = getNewMessages(incoming, existing)
  const messages = mergeMessages([...existing, ...incoming])
  await writeMessages(messages)
  void notifySlackMessages(slackEvents)
  res.json({ messages })
})

app.get('/api/users', async (_req, res) => {
  res.json({ users: await readUsers() })
})

app.post('/api/users/sync', async (req, res) => {
  const incoming = Array.isArray(req.body?.users) ? req.body.users : []
  const users = mergeUsers([...(await readUsers()), ...incoming])
  await writeUsers(users)
  res.json({ users })
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

async function readMessages() {
  try {
    const content = await fs.readFile(messagesFile, 'utf8')
    return mergeMessages(JSON.parse(content))
  } catch (error) {
    if (error.code !== 'ENOENT') console.error(error)
    return []
  }
}

async function writeMessages(messages) {
  await fs.mkdir(path.dirname(messagesFile), { recursive: true })
  await fs.writeFile(messagesFile, JSON.stringify(mergeMessages(messages), null, 2))
}

async function readUsers() {
  try {
    const content = await fs.readFile(usersFile, 'utf8')
    return mergeUsers(JSON.parse(content))
  } catch (error) {
    if (error.code !== 'ENOENT') console.error(error)
    return []
  }
}

async function writeUsers(users) {
  await fs.mkdir(path.dirname(usersFile), { recursive: true })
  await fs.writeFile(usersFile, JSON.stringify(mergeUsers(users), null, 2))
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

function normalizeMessage(message) {
  const now = new Date().toISOString()
  const createdAt = message.createdAt || now
  return {
    id: message.id || randomUUID(),
    type: ['aviso', 'mensaje', 'urgente'].includes(message.type) ? message.type : 'mensaje',
    author: String(message.author || 'Operacion'),
    text: String(message.text || '').trim(),
    createdAt,
    updatedAt: message.updatedAt || createdAt,
  }
}

function mergeMessages(messages) {
  const map = new Map()
  messages.map(normalizeMessage).filter((message) => message.text).forEach((message) => {
    const current = map.get(message.id)
    if (!current || new Date(message.updatedAt).getTime() >= new Date(current.updatedAt).getTime()) {
      map.set(message.id, message)
    }
  })
  return Array.from(map.values()).sort(
    (a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime(),
  )
}

function getNewMessages(incoming, existing) {
  const existingById = new Map(existing.map((message) => [message.id, message]))
  return incoming
    .map(normalizeMessage)
    .filter((message) => message.text && !existingById.has(message.id))
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

async function notifySlackMessages(messages) {
  const webhookUrl = process.env.SLACK_WEBHOOK_URL
  if (!webhookUrl || messages.length === 0) return

  const maxMessages = 5
  const publicUrl = process.env.PUBLIC_APP_URL || defaultPublicUrl
  const lines = messages.slice(0, maxMessages).map(formatSlackMessage)
  const extra = messages.length > maxMessages ? `\n...y ${messages.length - maxMessages} mensaje(s) mas.` : ''
  const text = [
    `MGA Bitacora Mina: ${messages.length} aviso(s) del centro de mensajes`,
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

function formatSlackMessage(message) {
  const typeLabel = {
    aviso: 'Aviso',
    mensaje: 'Mensaje',
    urgente: 'Urgente',
  }[message.type] || 'Mensaje'
  return `- ${typeLabel}: ${message.text} - ${message.author}`
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
    record.handoffNotes ? `Pase: ${shortText(record.handoffNotes, 90)}` : '',
  ].filter(Boolean)
  return `- ${actionLabel}: ${typeLabel} | ${details.join(' | ')}`
}

function normalizeUser(user) {
  const now = new Date().toISOString()
  const createdAt = user.createdAt || now
  return {
    id: user.id || randomUUID(),
    username: String(user.username || '').trim().toLowerCase(),
    displayName: String(user.displayName || user.username || 'Usuario').trim(),
    supervisorName: String(user.supervisorName || user.displayName || user.username || 'Supervisor').trim(),
    role: ['supervisor', 'administrador', 'gerencia'].includes(user.role) ? user.role : 'supervisor',
    pinHash: typeof user.pinHash === 'string' ? user.pinHash : undefined,
    passwordHash: typeof user.passwordHash === 'string' ? user.passwordHash : undefined,
    biometricEnabled: Boolean(user.biometricEnabled),
    active: user.active !== false,
    createdAt,
    updatedAt: user.updatedAt || createdAt,
    syncedAt: user.syncedAt,
    deletedAt: user.deletedAt,
  }
}

function mergeUsers(users) {
  const map = new Map()
  users.map(normalizeUser).filter((user) => user.username).forEach((user) => {
    const current = map.get(user.id)
    if (!current || new Date(user.updatedAt).getTime() >= new Date(current.updatedAt).getTime()) {
      map.set(user.id, user)
    }
  })
  return Array.from(map.values()).sort((a, b) => a.displayName.localeCompare(b.displayName))
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

function shortText(value, maxLength) {
  const clean = String(value || '').replace(/\s+/g, ' ').trim()
  return clean.length <= maxLength ? clean : `${clean.slice(0, maxLength - 3)}...`
}
