# QR Attendance (sem banco de dados)

Projeto de chamada por QR Code com validação de distância (≤ 50m), lista em tempo real e exportação Excel/Word. **Sem banco de dados**: os dados vivem só na memória e são descartados após exportar.

## Como rodar
1. Instale Node.js 18+
2. No terminal:
```bash
cd qr-attendance
npm install
npm run start
```
3. Abra http://localhost:3000/ no dispositivo do professor e permita a localização.
4. Os alunos escaneiam o QR e permitem a localização.

## Estrutura
- `server.js` — API + Socket.IO + exportação
- `public/index.html` — tela do professor
- `public/join.html` — tela do aluno
