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
  const records = mergeRecords([...existing, ...incoming])
  await writeRecords(records)
  res.json({ records })
})

app.put('/api/records/:id', async (req, res) => {
  const now = new Date().toISOString()
  const record = normalizeRecord({ ...req.body, id: req.params.id, updatedAt: now })
  const records = mergeRecords([...(await readRecords()), record])
  await writeRecords(records)
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
