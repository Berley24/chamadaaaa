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

// servir os arquivos estáticos do front
app.use(express.static(path.join(__dirname, 'public')));

// garantir que / sirva a tela do professor
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// memória (sem banco)
const sessions = new Map();
// { id: { name, createdAt, lat, lng, active, attendees: [{name, rgm, time, ip}] } }

function randomId(length = 8) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < length; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

// Haversine (m)
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

// criar chamada
app.post('/api/sessions', async (req, res) => {
  const { name, lat, lng } = req.body || {};
  if (!name || typeof lat !== 'number' || typeof lng !== 'number') {
    return res.status(400).json({ error: 'name, lat, lng são obrigatórios' });
  }
  const id = randomId();
  sessions.set(id, { name, createdAt: new Date(), lat, lng, active: true, attendees: [] });

  // usa host/proto corretos (Render/local/túnel)
  const proto = (req.headers['x-forwarded-proto'] || req.protocol).split(',')[0];
  const host = (req.headers['x-forwarded-host'] || req.get('host'));
  const joinUrl = `${proto}://${host}/join.html?id=${id}`;

  const qrPng = await QRCode.toDataURL(joinUrl);
  res.json({ id, name, joinUrl, qrPng });
});

// snapshot da sessão (professor usa)
app.get('/api/sessions/:id', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  res.json({ id: req.params.id, name: s.name, active: s.active, attendees: s.attendees });
});

// aluno envia presença
app.post('/api/sessions/:id/join', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  if (!s.active) return res.status(403).json({ error: 'Chamada encerrada' });

  const { name, rgm, lat, lng } = req.body || {};
  if (!name || !rgm || typeof lat !== 'number' || typeof lng !== 'number') {
    return res.status(400).json({ error: 'name, rgm, lat, lng são obrigatórios' });
  }

  if (distanceMeters(s.lat, s.lng, lat, lng) > 50) {
    return res.status(403).json({ error: 'Fora do raio permitido (50m)' });
  }

  // bloqueia duplicidade por RGM
  if (s.attendees.some(a => a.rgm.trim().toLowerCase() === String(rgm).trim().toLowerCase())) {
    return res.status(409).json({ error: 'RGM já registrado nesta chamada' });
  }

  const ip = req.headers['x-forwarded-for']?.split(',')[0]?.trim() || req.socket.remoteAddress;
  const attendee = { name, rgm: String(rgm), time: new Date().toISOString(), ip };
  s.attendees.push(attendee);

  io.to(`host:${req.params.id}`).emit('attendee:new', attendee);
  res.json({ ok: true });
});

// fechar chamada (para de aceitar envios)
app.post('/api/sessions/:id/close', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  s.active = false;
  res.json({ ok: true });
});

// exportar Excel (e apagar sessão)
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
    { header: 'IP', key: 'ip', width: 18 }
  ];
  s.attendees.forEach(a =>
    ws.addRow({ lesson: s.name, name: a.name, rgm: a.rgm, time: a.time, ip: a.ip })
  );

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="chamada_${req.params.id}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();

  sessions.delete(req.params.id); // apaga tudo após export
});

// exportar Word (e apagar sessão)
app.get('/api/sessions/:id/export.docx', async (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).send('Not found');

  const paragraphs = [
    new Paragraph({ children: [ new TextRun({ text: `Lista de Presença - ${s.name}`, bold: true, size: 28 }) ] }),
    new Paragraph({ children: [ new TextRun({ text: `Gerado em: ${new Date().toLocaleString()}` }) ] }),
    new Paragraph({ children: [ new TextRun({ text: '' }) ] })
  ];
  s.attendees.forEach((a, i) => {
    paragraphs.push(new Paragraph({ children: [ new TextRun({ text: `${i+1}. ${a.name} - ${a.rgm} - ${a.time}` }) ] }));
  });

  const doc = new Document({ sections: [{ properties: {}, children: paragraphs }] });
  const b = await Packer.toBuffer(doc);

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="chamada_${req.params.id}.docx"`);
  res.send(b);

  sessions.delete(req.params.id); // apaga tudo após export
});

// opcional: descartar sem baixar (purge)
app.post('/api/sessions/:id/purge', (req, res) => {
  if (!sessions.has(req.params.id)) return res.status(404).json({ error: 'Sessão não encontrada' });
  sessions.delete(req.params.id);
  res.json({ ok: true });
});

// socket.io (tempo real)
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
