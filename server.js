
import express from 'express';
import http from 'http';
import { Server as SocketIOServer } from 'socket.io';
import QRCode from 'qrcode';
import ExcelJS from 'exceljs';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import cookieParser from 'cookie-parser';
import morgan from 'morgan';
import cors from 'cors';

const app = express();
const server = http.createServer(app);
const io = new SocketIOServer(server, { cors: { origin: '*' } });

app.use(cors());
app.use(express.json());
app.use(cookieParser());
app.use(morgan('dev'));
app.use(express.static('public'));

// In-memory store (no database)
const sessions = new Map();
// session structure:
// id: { name, createdAt, lat, lng, active, attendees: [{name, rgm, time, ip}] }

function randomId(length = 8) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < length; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

// Haversine distance in meters
function distanceMeters(lat1, lon1, lat2, lon2) {
  function toRad(x){ return x*Math.PI/180; }
  const R = 6371000; // meters
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
            Math.sin(dLon/2) * Math.sin(dLon/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
}

// Create session
app.post('/api/sessions', async (req, res) => {
  const { name, lat, lng } = req.body || {};
  if (!name || typeof lat !== 'number' || typeof lng !== 'number') {
    return res.status(400).json({ error: 'name, lat, lng são obrigatórios' });
  }
  const id = randomId();
  sessions.set(id, { name, createdAt: new Date(), lat, lng, active: true, attendees: [] });
  const joinUrl = `https://creativity-rna-designation-ky.trycloudflare.com/join.html?id=${id}`;
  const qrPng = await QRCode.toDataURL(joinUrl);
  res.json({ id, name, joinUrl, qrPng });
});

// Get QR image (optional separate route)
app.get('/api/sessions/:id/qr.png', async (req, res) => {
  const id = req.params.id;
  if (!sessions.has(id)) return res.status(404).send('Not found');
  const joinUrl = `https://creativity-rna-designation-ky.trycloudflare.com/join.html?id=${id}`;
  const png = await QRCode.toBuffer(joinUrl);
  res.setHeader('Content-Type','image/png');
  res.send(png);
});

// Get session info (for host screen)
app.get('/api/sessions/:id', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  res.json({ id: req.params.id, name: s.name, active: s.active, attendees: s.attendees });
});

// Student submit
app.post('/api/sessions/:id/join', (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  if (!s.active) return res.status(403).json({ error: 'Chamada encerrada' });

  const { name, rgm, lat, lng } = req.body || {};
  if (!name || !rgm || typeof lat !== 'number' || typeof lng !== 'number') {
    return res.status(400).json({ error: 'name, rgm, lat, lng são obrigatórios' });
  }

  // distance check (<= 50m)
  const dist = distanceMeters(s.lat, s.lng, lat, lng);
  if (dist > 50) {
    return res.status(403).json({ error: 'Fora do raio permitido (50m)' });
  }

  // reject duplicate RGM
  if (s.attendees.some(a => a.rgm.trim().toLowerCase() === String(rgm).trim().toLowerCase())) {
    return res.status(409).json({ error: 'RGM já registrado nesta chamada' });
  }

  const ip = req.headers['x-forwarded-for']?.split(',')[0]?.trim() || req.socket.remoteAddress;
  const attendee = { name, rgm: String(rgm), time: new Date().toISOString(), ip };
  s.attendees.push(attendee);

  // notify host via socket
  io.to(`host:${req.params.id}`).emit('attendee:new', attendee);

  res.json({ ok: true });
});

// Close session
app.post('/api/sessions/:id/close', async (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).json({ error: 'Sessão não encontrada' });
  s.active = false;
  res.json({ ok: true });
});

// Export Excel
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

  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="chamada_${req.params.id}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();

  // purge after export
  sessions.delete(req.params.id);
});

// Export Word
app.get('/api/sessions/:id/export.docx', async (req, res) => {
  const s = sessions.get(req.params.id);
  if (!s) return res.status(404).send('Not found');
  const paragraphs = [
    new Paragraph({ children: [ new TextRun({ text: `Lista de Presença - ${s.name}`, bold: true, size: 28 }) ] }),
    new Paragraph({ children: [ new TextRun({ text: `Gerado em: ${new Date().toLocaleString()}` }) ] }),
    new Paragraph({ children: [ new TextRun({ text: '' }) ] }),
  ];
  s.attendees.forEach((a, i) => {
    paragraphs.push(new Paragraph({ children: [ new TextRun({ text: `${i+1}. ${a.name} - ${a.rgm} - ${a.time}` }) ] }));
  });
  const doc = new Document({ sections: [{ properties: {}, children: paragraphs }] });
  const b = await Packer.toBuffer(doc);
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename="chamada_${req.params.id}.docx"`);
  res.send(b);

  // purge after export
  sessions.delete(req.params.id);
});

// Socket.io
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
