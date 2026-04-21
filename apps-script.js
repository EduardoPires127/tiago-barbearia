// ================================================================
//  TIAGO BARBEARIA — Google Apps Script
//  Agendamento · E-mail confirmação/cancelamento · Lembrete 2h antes
//  Gratuito — usa Google Calendar + Gmail
// ================================================================

// ⚙️ CONFIGURAÇÃO — preencha os campos abaixo antes de publicar
const CONFIG = {

  // ── E-mail do Tiago (recebe notificações e avisa o barbeiro) ──
  emailBarbearia: 'tiagobarbearia.ofc@gmail.com',

  // ── Calendário do Tiago ────────────────────────────────────
  // Use o e-mail da conta Google do Tiago (ou o ID de um calendário específico)
  calendarioId: 'tiagobarbearia.ofc@gmail.com',

  // ── URL pública deste Apps Script (gerada após publicar) ──
  // Apps Script > Implantar > Nova implantação > App da Web
  urlScript: 'https://script.google.com/macros/s/AKfycbyVZABWd7ODGC8NObfMwPQNVFQWBOURGo7Ogc-6FoOvCIt6n7wpD1G4PmyUk6ZpDoji/exec',

  // ── URL do formulário (usada nos e-mails de cancelamento) ──
  urlFormulario: 'https://eduardopires127.github.io/tiago-barbearia/tiago-barbearia.html',

  // ── Planilha de registros ─────────────────────────────────
  // Crie uma planilha Google, copie o ID da URL e cole aqui
  // URL exemplo: docs.google.com/spreadsheets/d/SEU_ID_AQUI/edit
  planilhaId: '1Lee6J0Z2GEPQzkg-T_93pSND34utCldBZstDCO4zLFw',

  // ── Fuso horário ─────────────────────────────────────────
  timezone: 'America/Sao_Paulo',
};

// ================================================================
//  ROTAS GET
// ================================================================

function doGet(e) {
  const p = e.parameter;

  // Formulário buscando horários ocupados
  if (p.action === 'getEvents') {
    return getEvents(p.timeMin, p.timeMax);
  }

  // Cliente clicou em "Confirmar presença" no e-mail
  if (p.action === 'confirm' && p.token) {
    return confirmarViaLink(p.token);
  }

  // Cliente clicou em "Cancelar agendamento" no e-mail
  if (p.action === 'cancel' && p.token) {
    return cancelarViaLink(p.token);
  }

  // Busca agendamentos por telefone (para cancelamento pelo chat)
  if (p.action === 'buscarPorTelefone' && p.tel) {
    return buscarPorTelefone(p.tel);
  }

  // Cancelamento via chat (retorna JSON)
  if (p.action === 'cancelarJson' && p.token) {
    return cancelarJson(p.token);
  }

  if (p.action === 'getAgenda') return getAgenda(p.timeMin, p.timeMax);
  if (p.action === 'getHorarios') return getHorarios();

  // Formulário criando agendamento via GET (contorna CORS redirect do POST)
  if (p.action === 'createEvent') {
    try {
      return createEvent({
        nome:     p.nome,
        tel:      p.tel,
        email:    p.email,
        date:     p.date,
        time:     p.time,
        duration: Number(p.duration),
        servicos: p.servicos,
        valor:    p.valor,
      });
    } catch (err) {
      return jsonResp({ status: 'error', message: err.toString() });
    }
  }

  return ContentService.createTextOutput('OK');
}

// ================================================================
//  ROTAS POST
// ================================================================

function doPost(e) {
  let payload = {};
  try { payload = JSON.parse(e.postData.contents); } catch(_) {}

  if (payload.action === 'salvarHorarios') return salvarHorarios(payload);
  if (payload.action === 'salvarExcecao')  return salvarExcecao(payload);
  if (payload.action === 'removerExcecao') return removerExcecao(payload);

  return jsonResp({ status: 'error', message: 'Ação desconhecida' });
}

// ================================================================
//  BUSCAR EVENTOS (para o formulário bloquear horários ocupados)
// ================================================================

function getEvents(timeMin, timeMax) {
  try {
    const cal    = CalendarApp.getCalendarById(CONFIG.calendarioId);
    const events = cal.getEvents(new Date(timeMin), new Date(timeMax));
    const result = events.map(ev => {
      const startMs = ev.getStartTime().getTime();
      const endMs   = ev.getEndTime().getTime();
      const dur     = Math.round((endMs - startMs) / 60000);
      const dt      = ev.getStartTime();
      const hh      = String(dt.getHours()).padStart(2,'0');
      const mm      = String(dt.getMinutes()).padStart(2,'0');
      return { time: hh + ':' + mm, duration: dur };
    });
    return jsonResp({ status: 'ok', events: result });
  } catch (err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// ================================================================
//  CRIAR EVENTO + E-MAIL DE CONFIRMAÇÃO
// ================================================================

function createEvent(p) {
  try {
    const cal = CalendarApp.getCalendarById(CONFIG.calendarioId);

    const [year, month, day] = p.date.split('-').map(Number);
    const [hour, min]        = p.time.split(':').map(Number);
    const startTime = new Date(year, month - 1, day, hour, min, 0);
    const endTime   = new Date(startTime.getTime() + p.duration * 60000);

    // Verifica conflito de último minuto
    if (cal.getEvents(startTime, endTime).length > 0) {
      return jsonResp({ status: 'error', message: 'Este horário acabou de ser reservado. Escolha outro.' });
    }

    // Token único para links de confirmação/cancelamento
    const token = Utilities.getUuid().replace(/-/g, '');

    const description =
      `👤 Cliente: ${p.nome}\n` +
      `📱 Telefone: ${p.tel}\n` +
      `📧 E-mail: ${p.email}\n` +
      `✂️ Serviço(s): ${p.servicos}\n` +
      `💰 Valor: ${p.valor}\n` +
      `🔑 TOKEN: ${token}`;

    cal.createEvent(
      `✂️ ${p.nome} — Tiago Barbearia`,
      startTime,
      endTime,
      { description }
    );

    const dataFmt        = formatDateBR(startTime);
    const linkCancelar   = `${CONFIG.urlScript}?action=cancel&token=${token}`;
    const linkConfirmar  = `${CONFIG.urlScript}?action=confirm&token=${token}`;

    enviarEmailConfirmacao(p, dataFmt, linkCancelar);
    registrarAgendamento(p, token);

    return jsonResp({ status: 'ok' });
  } catch (err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// ================================================================
//  LEMBRETE AUTOMÁTICO 2H ANTES
//  Configure um Trigger em: Projeto > Gatilhos > lembreteAutomatico
//  Tipo: baseado em tempo → a cada hora
// ================================================================

function lembreteAutomatico() {
  const agora    = new Date();
  const em2h     = new Date(agora.getTime() + 2 * 60 * 60 * 1000);
  const margem   = 5 * 60 * 1000; // ±5 minutos de tolerância
  const buscaIni = new Date(em2h.getTime() - margem);
  const buscaFim = new Date(em2h.getTime() + margem);

  const cal    = CalendarApp.getCalendarById(CONFIG.calendarioId);
  const events = cal.getEvents(buscaIni, buscaFim);

  events.forEach(ev => {
    const desc = ev.getDescription();
    if (!desc || !desc.includes('TOKEN:')) return;

    const nome     = extrairDado(desc, 'Cliente');
    const email    = extrairDado(desc, 'E-mail');
    const tel      = extrairDado(desc, 'Telefone');
    const servicos = extrairDado(desc, 'Serviço(s)');
    const valor    = extrairDado(desc, 'Valor');
    const token    = extrairDado(desc, 'TOKEN');
    const horario  = Utilities.formatDate(ev.getStartTime(), CONFIG.timezone, 'HH:mm');
    const dataFmt  = formatDateBR(ev.getStartTime());

    const linkConfirmar = `${CONFIG.urlScript}?action=confirm&token=${token}`;
    const linkCancelar  = `${CONFIG.urlScript}?action=cancel&token=${token}`;

    // Lembrete por e-mail para o cliente
    const htmlCliente = emailTemplate(`
      <h2>Seu horário é em 2 horas! ⏰</h2>
      <p>Olá, <strong>${nome}</strong>! Não esqueça do seu agendamento:</p>
      <table>
        <tr><td>Serviço(s)</td><td>${servicos}</td></tr>
        <tr><td>Data</td><td>${dataFmt}</td></tr>
        <tr><td>Horário</td><td>${horario}</td></tr>
        <tr><td>Valor</td><td>${valor}</td></tr>
      </table>
      <p>Confirme sua presença ou cancele se necessário:</p>
      <div class="btn-row">
        <a href="${linkConfirmar}" class="btn-green">✅ Confirmar presença</a>
        <a href="${linkCancelar}"  class="btn-red">❌ Cancelar agendamento</a>
      </div>
      <p class="note">Caso não responda, seu agendamento permanece confirmado. Te esperamos! 💈</p>
    `);
    GmailApp.sendEmail(email, '⏰ Lembrete: seu horário é em 2 horas — Tiago Barbearia', '', { htmlBody: htmlCliente });
  });
}

// ================================================================
//  CONFIRMAR VIA LINK
// ================================================================

function confirmarViaLink(token) {
  let nome = '', horario = '', data = '';
  try {
    const ev = buscarEventoPorToken(token);
    if (ev) {
      nome    = extrairDado(ev.getDescription(), 'Cliente');
      horario = Utilities.formatDate(ev.getStartTime(), CONFIG.timezone, 'HH:mm');
      data    = formatDateBR(ev.getStartTime());
    }
  } catch(_) {}

  return HtmlService.createHtmlOutput(paginaResultado(
    '✅ Presença confirmada!',
    `${nome ? `Ótimo, <strong>${nome}</strong>! ` : ''}Te esperamos às <strong>${horario}</strong> em ${data}. Até logo! 💈`,
    'success'
  ));
}

// ================================================================
//  CANCELAR VIA LINK
// ================================================================

function cancelarViaLink(token) {
  try {
    const ev = buscarEventoPorToken(token);

    if (!ev) {
      return HtmlService.createHtmlOutput(paginaResultado(
        '⚠️ Não encontrado',
        'Este agendamento já foi cancelado ou o link não é mais válido.',
        'warning'
      ));
    }

    const desc     = ev.getDescription();
    const nome     = extrairDado(desc, 'Cliente');
    const tel      = extrairDado(desc, 'Telefone');
    const email    = extrairDado(desc, 'E-mail');
    const servicos = extrairDado(desc, 'Serviço(s)');
    const valor    = extrairDado(desc, 'Valor');
    const horario  = Utilities.formatDate(ev.getStartTime(), CONFIG.timezone, 'HH:mm');
    const data     = formatDateBR(ev.getStartTime());
    const tokenEv  = extrairDado(desc, 'TOKEN');

    ev.deleteEvent();
    enviarEmailCancelamento({ nome, tel, email, servicos, data, horario });
    registrarCancelamento(nome, tel, email, servicos, valor, data, horario, tokenEv);

    return HtmlService.createHtmlOutput(paginaResultado(
      '❌ Agendamento cancelado',
      `Tudo certo, <strong>${nome}</strong>! Seu agendamento de <strong>${servicos}</strong> em <strong>${data} às ${horario}</strong> foi cancelado e o horário foi liberado.`,
      'error'
    ));

  } catch (err) {
    Logger.log('cancelarViaLink error: ' + err);
    return HtmlService.createHtmlOutput(paginaResultado(
      '❌ Erro',
      'Não foi possível cancelar. Entre em contato com a barbearia pelo WhatsApp.',
      'error'
    ));
  }
}

// ================================================================
//  HELPER — busca evento pelo token nos próximos 60 dias
// ================================================================

function buscarEventoPorToken(token) {
  const now = new Date();
  const max = new Date(now.getTime() + 60 * 24 * 60 * 60 * 1000);
  const cal    = CalendarApp.getCalendarById(CONFIG.calendarioId);
  const events = cal.getEvents(now, max);
  for (const ev of events) {
    if (ev.getDescription().includes(`TOKEN: ${token}`)) return ev;
  }
  return null;
}

// ================================================================
//  E-MAILS
// ================================================================

function enviarEmailConfirmacao(p, dataFmt, linkCancelar) {
  // E-mail de confirmação para o cliente
  const htmlCliente = emailTemplate(`
    <h2>Agendamento confirmado! ✂️</h2>
    <p>Olá, <strong>${p.nome}</strong>! Seu horário foi reservado com sucesso.</p>
    <table>
      <tr><td>Serviço(s)</td><td>${p.servicos}</td></tr>
      <tr><td>Data</td><td>${dataFmt}</td></tr>
      <tr><td>Horário</td><td>${p.time}</td></tr>
      <tr><td>Valor</td><td>${p.valor}</td></tr>
    </table>
    <p>Você receberá um <strong>lembrete por e-mail 2 horas antes</strong> do seu horário.</p>
    <p>Precisa cancelar? Clique no botão abaixo:</p>
    <a href="${linkCancelar}" class="btn-red">❌ Cancelar agendamento</a>
    <p class="note" style="margin-top:8px;font-size:11px;color:#888;word-break:break-all">${linkCancelar}</p>
  `);
  GmailApp.sendEmail(p.email, '✅ Agendamento confirmado — Tiago Barbearia', '', { htmlBody: htmlCliente });

  // Notificação para o Tiago
  const telLimpo   = p.tel.replace(/\D/g, '');
  const wppCliente = `https://wa.me/55${telLimpo}`;
  const htmlTiago  = emailTemplate(`
    <h2>📅 Novo agendamento!</h2>
    <table>
      <tr><td>👤 Cliente</td><td><strong>${p.nome}</strong></td></tr>
      <tr><td>📱 WhatsApp</td><td><a href="${wppCliente}" style="color:#D4913A">${p.tel}</a></td></tr>
      <tr><td>📧 E-mail</td><td>${p.email}</td></tr>
      <tr><td>✂️ Serviço(s)</td><td>${p.servicos}</td></tr>
      <tr><td>📆 Data</td><td>${dataFmt}</td></tr>
      <tr><td>⏰ Horário</td><td><strong>${p.time}</strong></td></tr>
      <tr><td>💰 Valor</td><td><strong>${p.valor}</strong></td></tr>
    </table>
    <p>O horário foi registrado no calendário automaticamente.</p>
    <a href="${linkCancelar}" class="btn-red">❌ Cancelar este agendamento</a>
  `);
  GmailApp.sendEmail(CONFIG.emailBarbearia, `📅 Novo agendamento: ${p.nome} — ${dataFmt} às ${p.time}`, '', { htmlBody: htmlTiago });
}

function enviarEmailCancelamento(d) {
  // E-mail ao cliente (se tiver e-mail cadastrado)
  if (d.email) {
    const htmlCliente = emailTemplate(`
      <h2>Agendamento cancelado</h2>
      <p>Olá, <strong>${d.nome}</strong>. Seu agendamento foi cancelado:</p>
      <table>
        <tr><td>Serviço(s)</td><td>${d.servicos}</td></tr>
        <tr><td>Data</td><td>${d.data}</td></tr>
        <tr><td>Horário</td><td>${d.horario}</td></tr>
      </table>
      <p>O horário foi liberado. Para reagendar:</p>
      <a href="${CONFIG.urlFormulario}" class="btn-green">📅 Reagendar agora</a>
      <p class="note">Esperamos te ver em breve! 💈</p>
    `);
    GmailApp.sendEmail(d.email, '❌ Agendamento cancelado — Tiago Barbearia', '', { htmlBody: htmlCliente });
  }

  // Notificação para o Tiago
  const telLimpo   = (d.tel || '').replace(/\D/g, '');
  const wppCliente = telLimpo ? `https://wa.me/55${telLimpo}` : '';
  const htmlTiago  = emailTemplate(`
    <h2>❌ Agendamento cancelado</h2>
    <table>
      <tr><td>👤 Cliente</td><td><strong>${d.nome}</strong></td></tr>
      ${wppCliente ? `<tr><td>📱 WhatsApp</td><td><a href="${wppCliente}" style="color:#D4913A">${d.tel}</a></td></tr>` : ''}
      <tr><td>✂️ Serviço(s)</td><td>${d.servicos}</td></tr>
      <tr><td>📆 Data</td><td>${d.data}</td></tr>
      <tr><td>⏰ Horário</td><td>${d.horario}</td></tr>
    </table>
    <p>O horário foi liberado automaticamente no calendário.</p>
  `);
  GmailApp.sendEmail(CONFIG.emailBarbearia, `❌ Cancelamento: ${d.nome} — ${d.data} às ${d.horario}`, '', { htmlBody: htmlTiago });
}

// ================================================================
//  REGISTRAR NA PLANILHA
// ================================================================

function registrarAgendamento(p, token) {
  try {
    const ss     = SpreadsheetApp.openById(CONFIG.planilhaId);
    let sheet    = ss.getSheetByName('Agendamentos');
    if (!sheet) {
      sheet = ss.insertSheet('Agendamentos');
      sheet.appendRow(['Data/Hora Registro','Nome','Telefone','E-mail','Serviço(s)','Valor','Data','Horário','Status','Token']);
    }
    sheet.appendRow([
      new Date(),
      p.nome, p.tel, p.email,
      p.servicos, p.valor,
      p.date, p.time,
      'Agendado', token
    ]);
  } catch(e) {
    Logger.log('registrarAgendamento error: ' + e);
  }
}

function registrarCancelamento(nome, tel, email, servicos, valor, data, horario, token) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.planilhaId);
    let sheet   = ss.getSheetByName('Agendamentos');
    if (!sheet) return;
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][9] === token) {
        sheet.getRange(i + 1, 9).setValue('Cancelado');
        break;
      }
    }
  } catch(e) {
    Logger.log('registrarCancelamento error: ' + e);
  }
}

// ================================================================
//  BUSCAR AGENDAMENTO POR TELEFONE (para cancelamento pelo chat)
// ================================================================

function buscarPorTelefone(tel) {
  try {
    const telLimpo = tel.replace(/\D/g, '');
    const now      = new Date();
    const max      = new Date(now.getTime() + 60 * 24 * 60 * 60 * 1000);
    const cal      = CalendarApp.getCalendarById(CONFIG.calendarioId);
    const events   = cal.getEvents(now, max);
    const result   = [];

    for (const ev of events) {
      const desc = ev.getDescription() || '';
      if (!desc.includes('TOKEN:')) continue;
      const telEv = extrairDado(desc, 'Telefone').replace(/\D/g, '');
      // compara os últimos 9 dígitos para flexibilidade (com/sem DDI)
      if (!telEv || !telLimpo) continue;
      const match = telEv.slice(-9) === telLimpo.slice(-9);
      if (!match) continue;
      result.push({
        nome:     extrairDado(desc, 'Cliente'),
        servicos: extrairDado(desc, 'Serviço(s)'),
        valor:    extrairDado(desc, 'Valor'),
        horario:  Utilities.formatDate(ev.getStartTime(), CONFIG.timezone, 'HH:mm'),
        data:     formatDateBR(ev.getStartTime()),
        token:    extrairDado(desc, 'TOKEN'),
      });
    }
    return jsonResp({ status: 'ok', agendamentos: result });
  } catch(err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// ================================================================
//  CANCELAR VIA CHAT (retorna JSON para o chatbot)
// ================================================================

function cancelarJson(token) {
  try {
    const ev = buscarEventoPorToken(token);
    if (!ev) return jsonResp({ status: 'not_found' });

    const desc     = ev.getDescription();
    const nome     = extrairDado(desc, 'Cliente');
    const tel      = extrairDado(desc, 'Telefone');
    const email    = extrairDado(desc, 'E-mail');
    const servicos = extrairDado(desc, 'Serviço(s)');
    const valor    = extrairDado(desc, 'Valor');
    const horario  = Utilities.formatDate(ev.getStartTime(), CONFIG.timezone, 'HH:mm');
    const data     = formatDateBR(ev.getStartTime());
    const tokenEv  = extrairDado(desc, 'TOKEN');

    ev.deleteEvent();
    enviarEmailCancelamento({ nome, tel, email, servicos, data, horario });
    registrarCancelamento(nome, tel, email, servicos, valor, data, horario, tokenEv);

    return jsonResp({ status: 'ok', nome, servicos, data, horario });
  } catch(err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// ================================================================
//  AGENDA ADMIN — getAgenda
// ================================================================

// getAgenda — retorna todos eventos de um período
function getAgenda(timeMin, timeMax) {
  try {
    const cal    = CalendarApp.getCalendarById(CONFIG.calendarioId);
    const events = cal.getEvents(new Date(timeMin), new Date(timeMax));
    const result = events.map(ev => {
      const desc  = ev.getDescription() || '';
      const token = desc.includes('TOKEN:') ? extrairDado(desc, 'TOKEN') : '';
      return {
        data:     Utilities.formatDate(ev.getStartTime(), CONFIG.timezone, 'yyyy-MM-dd'),
        horario:  Utilities.formatDate(ev.getStartTime(), CONFIG.timezone, 'HH:mm'),
        horFim:   Utilities.formatDate(ev.getEndTime(),   CONFIG.timezone, 'HH:mm'),
        nome:     desc.includes('TOKEN:') ? extrairDado(desc, 'Cliente') : ev.getTitle().replace(/^✂️\s*/,''),
        servicos: extrairDado(desc, 'Serviço(s)') || '—',
        valor:    extrairDado(desc, 'Valor') || '',
        token,
      };
    });
    result.sort((a,b) => a.horario.localeCompare(b.horario));
    return jsonResp({ status: 'ok', events: result });
  } catch(err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// ================================================================
//  AGENDA ADMIN — getHorarios / salvarHorarios / exceções
// ================================================================

// getHorarios — lê config do PropertiesService
function getHorarios() {
  try {
    const props = PropertiesService.getScriptProperties();
    const json  = props.getProperty('horariosConfig');
    if (json) return jsonResp(Object.assign({ status: 'ok' }, JSON.parse(json)));
    // config padrão
    return jsonResp({
      status: 'ok',
      padrao: {
        0: { ativo: false },
        1: { ativo: true, periodos: [{inicio:9,fim:12},{inicio:14,fim:19}] },
        2: { ativo: true, periodos: [{inicio:9,fim:12},{inicio:14,fim:19}] },
        3: { ativo: true, periodos: [{inicio:9,fim:12},{inicio:14,fim:19}] },
        4: { ativo: true, periodos: [{inicio:9,fim:12},{inicio:14,fim:19}] },
        5: { ativo: true, periodos: [{inicio:9,fim:12},{inicio:14,fim:19}] },
        6: { ativo: true, periodos: [{inicio:9,fim:19}] },
      },
      excecoes: [],
    });
  } catch(err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// salvarHorarios
function salvarHorarios(payload) {
  try {
    const props = PropertiesService.getScriptProperties();
    const config = { padrao: payload.padrao, excecoes: payload.excecoes || [] };
    props.setProperty('horariosConfig', JSON.stringify(config));
    return jsonResp({ status: 'ok' });
  } catch(err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// salvarExcecao
function salvarExcecao(payload) {
  try {
    const props  = PropertiesService.getScriptProperties();
    const json   = props.getProperty('horariosConfig');
    const config = json ? JSON.parse(json) : { padrao: {}, excecoes: [] };
    if (!config.excecoes) config.excecoes = [];
    config.excecoes = config.excecoes.filter(e => e.data !== payload.data);
    config.excecoes.push({ data: payload.data, ativo: payload.ativo, motivo: payload.motivo || '' });
    props.setProperty('horariosConfig', JSON.stringify(config));
    return jsonResp({ status: 'ok' });
  } catch(err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// removerExcecao
function removerExcecao(payload) {
  try {
    const props  = PropertiesService.getScriptProperties();
    const json   = props.getProperty('horariosConfig');
    const config = json ? JSON.parse(json) : { padrao: {}, excecoes: [] };
    config.excecoes = (config.excecoes || []).filter(e => e.data !== payload.data);
    props.setProperty('horariosConfig', JSON.stringify(config));
    return jsonResp({ status: 'ok' });
  } catch(err) {
    return jsonResp({ status: 'error', message: err.toString() });
  }
}

// ================================================================
//  UTILITÁRIOS
// ================================================================

function extrairDado(desc, campo) {
  const regex = new RegExp(campo + ':\\s*(.+)');
  const match = desc.match(regex);
  return match ? match[1].trim() : '';
}

function formatDateBR(d) {
  const dias   = ['Domingo','Segunda-feira','Terça-feira','Quarta-feira','Quinta-feira','Sexta-feira','Sábado'];
  const meses  = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  return `${dias[d.getDay()]}, ${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`;
}

function jsonResp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================================================
//  TEMPLATE DE E-MAIL
// ================================================================

function emailTemplate(content) {
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<style>
  body { font-family: 'Helvetica Neue', Arial, sans-serif; background: #0D0C09; color: #F0E8D8; margin: 0; padding: 0; }
  .wrapper { max-width: 560px; margin: 0 auto; padding: 32px 16px 48px; }
  .header { text-align: center; margin-bottom: 28px; }
  .header h1 { font-size: 22px; color: #D4913A; letter-spacing: .06em; margin: 0; }
  .header p  { font-size: 12px; color: #7D7464; margin: 4px 0 0; letter-spacing: .1em; text-transform: uppercase; }
  .divider { width: 40px; height: 1px; background: #D4913A; margin: 10px auto; opacity: .5; }
  .body { background: #181610; border: 1px solid #2E2B22; border-radius: 12px; padding: 28px 24px; }
  .body h2 { font-size: 18px; color: #F0E8D8; margin: 0 0 10px; }
  .body p  { font-size: 14px; color: #BFB59E; line-height: 1.6; margin: 0 0 14px; }
  table { width: 100%; border-collapse: collapse; margin: 14px 0; }
  td { padding: 8px 10px; font-size: 14px; border-bottom: 1px solid #2E2B22; }
  td:first-child { color: #7D7464; width: 38%; }
  td:last-child  { color: #F0E8D8; font-weight: 500; }
  .btn-green, .btn-red { display: inline-block; padding: 11px 24px; border-radius: 8px; text-decoration: none; font-size: 14px; font-weight: 500; margin: 4px 6px 4px 0; }
  .btn-green { background: #3A9E6A; color: #fff; }
  .btn-red   { background: #B84040; color: #fff; }
  .btn-row   { margin: 16px 0 8px; }
  .note { font-size: 12px; color: #7D7464; line-height: 1.6; margin: 10px 0 0; }
  .footer { text-align: center; margin-top: 20px; font-size: 11px; color: #7D7464; }
</style>
</head>
<body>
<div class="wrapper">
  <div class="header">
    <h1>✂️ Tiago Barbearia</h1>
    <div class="divider"></div>
    <p>Agendamento Online</p>
  </div>
  <div class="body">
    ${content}
  </div>
  <div class="footer">Tiago Barbearia &nbsp;|&nbsp; Responda ao e-mail em caso de dúvidas</div>
</div>
</body>
</html>`;
}

// ================================================================
//  PÁGINA DE RESULTADO (para links confirm/cancel)
// ================================================================

function paginaResultado(titulo, mensagem, tipo) {
  const cores = { success: '#3A9E6A', error: '#B84040', warning: '#D4913A' };
  const cor   = cores[tipo] || '#D4913A';
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${titulo} — Tiago Barbearia</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Helvetica Neue', Arial, sans-serif; background: #0D0C09; color: #F0E8D8; display: flex; align-items: center; justify-content: center; min-height: 100vh; padding: 24px; }
  .card { background: #181610; border: 1px solid #2E2B22; border-radius: 16px; padding: 36px 28px; max-width: 420px; width: 100%; text-align: center; }
  .icon { font-size: 42px; margin-bottom: 14px; }
  h1 { font-size: 20px; color: ${cor}; margin-bottom: 12px; }
  p  { font-size: 14px; color: #BFB59E; line-height: 1.7; }
  a  { display: inline-block; margin-top: 22px; background: #D4913A; color: #0D0C09; text-decoration: none; padding: 12px 28px; border-radius: 8px; font-weight: 500; font-size: 14px; }
</style>
</head>
<body>
<div class="card">
  <div class="icon">${tipo === 'success' ? '✅' : tipo === 'error' ? '❌' : '⚠️'}</div>
  <h1>${titulo}</h1>
  <p>${mensagem}</p>
  <a href="${CONFIG.urlFormulario}">Agendar novamente</a>
</div>
</body>
</html>`;
}
