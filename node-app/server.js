const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const fs = require('fs');
const path = require('path');

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
  cors: { origin: "*" }
});

const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'quiz-data.json');

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Dados padrão com controle de liberação de envio e senha mestra
const defaultData = {
  settings: {
    lessonTitle: "Simulado Geral",
    backgroundUrl: "",
    allowSubmission: false, // Inicia em modo de espera (envio bloqueado)
    adminPassword: "prof123"
  },
  questions: [
    {
      id: 1,
      title: "Qual é a capital do Brasil?",
      options: { A: "Rio de Janeiro", B: "São Paulo", C: "Brasília", D: "Salvador" },
      correct: "C",
      visible: true
    }
  ],
  responses: []
};

function loadData() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const parsed = JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
      // Garantir que novas propriedades existam
      if (!parsed.settings) parsed.settings = {};
      if (parsed.settings.allowSubmission === undefined) parsed.settings.allowSubmission = false;
      if (!parsed.settings.adminPassword) parsed.settings.adminPassword = "prof123";
      return parsed;
    }
  } catch (err) {
    console.error("Erro ao carregar dados:", err);
  }
  saveData(defaultData);
  return defaultData;
}

function saveData(data) {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (err) {
    console.error("Erro ao salvar dados:", err);
  }
}

let db = loadData();

io.on('connection', (socket) => {
  console.log(`[+] Conectado: ${socket.id}`);

  // Envia estado inicial para o cliente
  socket.emit('init_state', {
    settings: {
      lessonTitle: db.settings.lessonTitle || "Simulado Geral",
      backgroundUrl: db.settings.backgroundUrl || "",
      allowSubmission: !!db.settings.allowSubmission
    },
    questions: db.questions,
    responses: db.responses
  });

  // 1. Salvar ou criar questão
  socket.on('save_question', (qData) => {
    if (qData.id) {
      const idx = db.questions.findIndex(x => x.id === qData.id);
      if (idx !== -1) {
        db.questions[idx] = qData;
      } else {
        db.questions.push(qData);
      }
    } else {
      qData.id = Date.now();
      db.questions.push(qData);
    }
    saveData(db);
    io.emit('questions_updated', db.questions);
    console.log(`[!] Questão salva: "${qData.title}"`);
  });

  // 2. Liberar ou Ocultar uma questão individual
  socket.on('toggle_question', (id) => {
    const q = db.questions.find(x => x.id === id);
    if (q) {
      q.visible = !q.visible;
      saveData(db);
      io.emit('questions_updated', db.questions);
      console.log(`[!] Questão ${id} visibilidade: ${q.visible}`);
    }
  });

  // 3. Liberar ou Ocultar Todas as questões
  socket.on('toggle_all_questions', (status) => {
    db.questions.forEach(q => q.visible = status);
    saveData(db);
    io.emit('questions_updated', db.questions);
    console.log(`[!] Todas questões visibilidade: ${status}`);
  });

  // 4. Excluir questão
  socket.on('delete_question', (id) => {
    db.questions = db.questions.filter(x => x.id !== id);
    saveData(db);
    io.emit('questions_updated', db.questions);
    console.log(`[!] Questão ${id} removida`);
  });

  // 5. Atualizar configurações (Título ou Imagem de Fundo)
  socket.on('update_settings', (newSettings) => {
    db.settings = { ...db.settings, ...newSettings };
    saveData(db);
    io.emit('settings_updated', {
      lessonTitle: db.settings.lessonTitle,
      backgroundUrl: db.settings.backgroundUrl,
      allowSubmission: !!db.settings.allowSubmission
    });
    console.log(`[!] Configurações atualizadas:`, db.settings);
  });

  // 6. Controle do Professor: Liberar ou Bloquear o envio final de respostas
  socket.on('toggle_allow_submission', (allow) => {
    db.settings.allowSubmission = !!allow;
    saveData(db);
    io.emit('settings_updated', {
      lessonTitle: db.settings.lessonTitle,
      backgroundUrl: db.settings.backgroundUrl,
      allowSubmission: db.settings.allowSubmission
    });
    console.log(`[!] Liberação de envio dos alunos alterada para: ${db.settings.allowSubmission}`);
  });

  // 7. Validação de senha do painel
  socket.on('verify_admin_password', (pwd, callback) => {
    const valid = (pwd === (db.settings.adminPassword || "prof123"));
    if (typeof callback === 'function') callback(valid);
  });

  // 8. Alteração de senha do painel
  socket.on('change_admin_password', (newPwd, callback) => {
    if (newPwd && newPwd.trim().length > 0) {
      db.settings.adminPassword = newPwd.trim();
      saveData(db);
      if (typeof callback === 'function') callback(true);
      console.log(`[!] Nova senha de administrador configurada`);
    } else {
      if (typeof callback === 'function') callback(false);
    }
  });

  // 9. Aluno enviando respostas
  socket.on('submit_response', (payload) => {
    // Validação de segurança: só aceita se o professor tiver liberado
    if (!db.settings.allowSubmission) {
      socket.emit('submit_error', { message: "O envio de respostas ainda não foi autorizado pelo professor." });
      return;
    }

    const newEntry = {
      id: Date.now(),
      date: payload.date || new Date().toLocaleString('pt-BR'),
      nome: payload.nome || '',
      sobrenome: payload.sobrenome || '',
      acertos: payload.acertos || '0',
      respostas: payload.respostas || {}
    };

    db.responses.push(newEntry);
    saveData(db);

    io.emit('new_response', newEntry);
    socket.emit('submit_success', { ok: true });
    console.log(`[+] Resposta recebida de ${newEntry.nome} ${newEntry.sobrenome} (${newEntry.acertos})`);
  });

  // 10. Limpar respostas
  socket.on('clear_responses', () => {
    db.responses = [];
    saveData(db);
    io.emit('responses_cleared');
    console.log(`[!] Base de respostas limpa`);
  });

  socket.on('disconnect', () => {
    console.log(`[-] Desconectado: ${socket.id}`);
  });
});

server.listen(PORT, () => {
  console.log(`=================================================`);
  console.log(`🚀 Servidor Quiz Rodando na Porta ${PORT}`);
  console.log(`=================================================`);
});
