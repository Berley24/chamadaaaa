import express from 'express';
import http from 'http';
import { Server as SocketIOServer } from 'socket.io';
import QRCode from 'qrcode';
import ExcelJS from 'exceljs';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import cookieParser from 'cookie-parser';
import morgan from 'morgan';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';

const app = express();
const server = http.createServer(app);
const io = new SocketIOServer(server, { cors: { origin: '*' } });

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(cors());
app.use(express.json());
app.use(cookieParser());
app.use(morgan('dev'));

// rota de saúde
app.get('/api/health', (_req, res) => res.json({ ok: true }));

// log simples
app.use((req, _res, next) => {
  console.log(`[REQ] ${req.method} ${req.url}`);
  next();
});

// front
app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// memória (sem DB)
const sessions = new Map();
// { id: { name, createdAt, lat, lng, active, attendees: [{name, rgm, rgmKey, time, ip}] } }

function randomId(length = 8) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < length; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

// Haversine (metros)
function distanceMeters(lat1, lon1, lat2, lon2) {
  const toRad = (x) => (x * Math.PI) / 180;
  const R = 6371000;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function normRGM(v) {
  return String(v)
    .normalize('NFKC')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9]/g, '')
    .trim();
}

// ------- helpers para nome de arquivo -------
function safeFilename(s) {
  return String(s || '')
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')  // remove acentos
    .replace(/[^a-z0-9\s\-_]/gi, '')                    // mantém só letras/números/espaço/-/_
    .trim()
    .replace(/\s+/g, '-')                               // espaços -> hífen
    .toLowerCase();
}
function todayISO() {
  return new Date().toISOString().slice(0, 10); // AAAA-MM-DD
}

// criar chamada
app.post('/api/sessions', async (req, res) => {
  const { name, lat, lng } = req.body || {};
  if (!name || typeof lat !== 'number' || typeof lng !== 'number') {
    return res.status(400).json({ error: 'name, lat, lng são obrigatórios' });
  }
  const id = randomId();
  sessions.set(id, { name, createdAt: new Date(), lat, lng, active: true, attendees: [] });

  const proto = (req.headers['x-forwarded-proto'] || req.protocol).split(',')[0];
  const host  = (req.headers['x-forwarded-host'] || req.get('host'));
  const joinUrl = `${proto}://${host}/join.html?id=${id}`;

  const qrPng = await QRCode.toDataURL(joinUrl);
  res.json({ id, name, joinUrl, qrPng });
});

// snapshot (professor)
app.get('/api/sessions/:id', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  res.json({ id: req.params.id, name: s.name, active: s.active, attendees: s.attendees });
});

// QR como PNG (para reaparecer após refresh)
app.get('/api/sessions/:id/qr.png', async (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).send('Not found');
  const proto = (req.headers['x-forwarded-proto'] || req.protocol).split(',')[0];
  const host  = (req.headers['x-forwarded-host'] || req.get('host'));
  const joinUrl = `${proto}://${host}/join.html?id=${req.params.id}`;
  try {
    const buf = await QRCode.toBuffer(joinUrl, { type: 'png', width: 256, errorCorrectionLevel: 'M' });
    res.setHeader('Content-Type', 'image/png');
    res.end(buf);
  } catch {
    res.status(500).send('QR error');
  }
});

// debug opcional
app.get('/api/sessions/:id/debug', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  res.json(s);
});

// atualizar localização da sessão (professor recenter)
app.patch('/api/sessions/:id/location', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });

  const { lat, lng } = req.body || {};
  if (typeof lat !== 'number' || typeof lng !== 'number') {
    return res.status(400).json({ error: 'lat, lng são obrigatórios' });
  }
  s.lat = lat;
  s.lng = lng;
  console.log(`[RELOC] id=${req.params.id} prof=(${lat.toFixed(6)},${lng.toFixed(6)})`);
  return res.json({ ok: true });
});

// aluno entra (validação 100m + logs)
app.post('/api/sessions/:id/join', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  if (!s.active) return res.status(403).json({ error: 'Chamada encerrada' });

  const { name, rgm, lat, lng, acc } = req.body || {};
  if (!name || !rgm || typeof lat !== 'number' || typeof lng !== 'number') {
    return res.status(400).json({ error: 'name, rgm, lat, lng são obrigatórios' });
  }

  const dist = distanceMeters(s.lat, s.lng, lat, lng);
  console.log(
    `[JOIN] id=${req.params.id} name=${name} rgm=${rgm} dist=${Math.round(dist)}m acc=${acc ?? 'n/a'} ` +
    `prof=(${s.lat.toFixed(6)},${s.lng.toFixed(6)}) aluno=(${lat.toFixed(6)},${lng.toFixed(6)})`
  );

  // RAIO INEGOCIÁVEL: 100m
  if (dist > 100) {
    return res.status(403).json({ error: 'Fora do raio permitido (100m)' });
  }

  const rgmKey = normRGM(rgm);
  const ip = req.headers['x-forwarded-for']?.split(',')[0]?.trim() || req.socket.remoteAddress;
  const cookieFlag = req.cookies?.[`att_${req.params.id}`];

  if (cookieFlag) return res.status(409).json({ error: 'Este dispositivo já registrou presença nesta chamada.' });
  if (s.attendees.some(a => a.rgmKey === rgmKey)) return res.status(409).json({ error: 'RGM já registrado nesta chamada.' });
  if (s.attendees.some(a => a.ip === ip)) return res.status(409).json({ error: 'Este IP já registrou presença nesta chamada.' });

  const attendee = { name, rgm: String(rgm), rgmKey, time: new Date().toISOString(), ip };
  s.attendees.push(attendee);

  res.cookie(`att_${req.params.id}`, '1', { maxAge: 24 * 60 * 60 * 1000, httpOnly: false, sameSite: 'Lax' });

  io.to(`host:${req.params.id}`).emit('attendee:new', attendee);
  res.json({ ok: true });
});

// fechar chamada (bloqueia novas entradas)
app.post('/api/sessions/:id/close', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  s.active = false;
  res.json({ ok: true });
});

// exportar Excel (NÃO apaga mais)
app.get('/api/sessions/:id/export.xlsx', async (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).send('Not found');

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Chamada');
  ws.columns = [
    { header: 'Nome da Aula', key: 'lesson', width: 30 },
    { header: 'Nome', key: 'name', width: 25 },
    { header: 'RGM', key: 'rgm', width: 15 },
    { header: 'Data/Hora', key: 'time', width: 24 },
    { header: 'IP', key: 'ip', width: 18 },
  ];
  s.attendees.forEach(a => ws.addRow({ lesson: s.name, name: a.name, rgm: a.rgm, time: a.time, ip: a.ip }));

  const filename = `lista-de-presenca_${safeFilename(s.name)}_${todayISO()}.xlsx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

  const buffer = await wb.xlsx.writeBuffer();
  res.end(Buffer.from(buffer));
});

// exportar Word (NÃO apaga mais)
app.get('/api/sessions/:id/export.docx', async (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).send('Not found');

  const paragraphs = [
    new Paragraph({ children: [ new TextRun({ text: `Lista de Presença - ${s.name}`, bold: true, size: 28 }) ] }),
    new Paragraph({ children: [ new TextRun({ text: `Gerado em: ${new Date().toLocaleString('pt-BR')}` }) ] }),
    new Paragraph({ children: [ new TextRun({ text: '' }) ] }),
  ];
  s.attendees.forEach((a, i) => {
    paragraphs.push(new Paragraph({ children: [ new TextRun({ text: `${i+1}. ${a.name} - ${a.rgm} - ${a.time}` }) ] }));
  });

  const doc = new Document({ sections: [{ properties: {}, children: paragraphs }] });
  const b = await Packer.toBuffer(doc);

  const filename = `lista-de-presenca_${safeFilename(s.name)}_${todayISO()}.docx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.end(b);
});

// limpar/descartar tudo manualmente
app.post('/api/sessions/:id/purge', (req, res) => {
  if (!sessions.has(req.params.id)) return res.status(404).json({ error: 'Sessão não encontrada' });
  sessions.delete(req.params.id);
  res.json({ ok: true });
});

// socket
io.on('connection', (socket) => {
  socket.on('host:join', (sessionId) => {
    if (!sessions.has(sessionId)) return;
    socket.join(`host:${sessionId}`);
  });
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log('Servidor rodando na porta ' + PORT);
});
