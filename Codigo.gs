/**
 * Sorteio Cesta Cristófoli · Manicures que Inspiram (Dermomed)
 * Recebe os cadastros da landing page, grava na planilha
 * e alimenta a área da consultora (lista de participantes + registro do sorteio).
 *
 * 1. Troque a SENHA abaixo pela senha da consultora.
 * 2. Implantar > Nova implantação > Tipo: App da Web
 *    Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 * 3. Copie a URL que termina em /exec e cole nas duas páginas (index.html e sorteio.html).
 */

const SENHA = 'TROQUE-ESTA-SENHA';

const TZ = 'America/Sao_Paulo';
const ABA_CADASTROS = 'Cadastros';
const ABA_SORTEIO = 'Sorteio';
const CAB_CADASTROS = ['Data/hora', 'Nome', 'WhatsApp', 'E-mail', 'Instagram', 'Atuação', 'Cidade/UF',
  'Segue @oficialdermomed', 'Aceite regulamento e privacidade', 'Aceita comunicação'];
const CAB_SORTEIO = ['Data/hora', 'Posição', 'Nome', 'WhatsApp', 'E-mail', 'Instagram', 'Situação', 'Consultora'];

function doGet() {
  return json({ ok: true, servico: 'sorteio-cesta-cristofoli' });
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.acao === 'listar') return listar(p);
    if (p.acao === 'resultado') return resultado(p);
    return cadastrar(p);
  } catch (err) {
    return json({ ok: false, erro: 'Erro no servidor: ' + err.message });
  }
}

/* ---------- cadastro vindo da landing ---------- */
function cadastrar(p) {
  if (p.empresa) return json({ ok: true }); // armadilha anti-robô

  const nome = limpa(p.nome);
  const whats = soDigitos(p.whatsapp);
  const email = limpa(p.email).toLowerCase();
  const insta = limpa(p.instagram);

  if (nome.split(/\s+/).length < 2) return json({ ok: false, erro: 'Digite nome e sobrenome.' });
  if (whats.length < 10 || whats.length > 11) return json({ ok: false, erro: 'Confira o WhatsApp com DDD.' });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json({ ok: false, erro: 'Confira o e-mail.' });
  if (p.regulamento !== 'sim' || p.segue_instagram !== 'sim') {
    return json({ ok: false, erro: 'Aceite o regulamento e confirme que segue o @oficialdermomed.' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = aba(ABA_CADASTROS, CAB_CADASTROS);
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      const dados = sh.getRange(2, 3, n, 2).getValues(); // WhatsApp, E-mail
      const repetido = dados.some(function (r) {
        return soDigitos(r[0]) === whats || String(r[1]).trim().toLowerCase() === email;
      });
      if (repetido) return json({ ok: true, duplicado: true });
    }
    sh.appendRow([
      agora(), texto(nome), texto(formataFone(whats)), texto(email), texto(insta),
      texto(p.atuacao), texto(p.cidade), 'sim', 'sim', p.aceita_comunicacao === 'sim' ? 'sim' : 'não'
    ]);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

/* ---------- área da consultora ---------- */
function listar(p) {
  if (p.senha !== SENHA) return json({ ok: false, erro: 'Senha incorreta.' });
  const sh = aba(ABA_CADASTROS, CAB_CADASTROS);
  const n = sh.getLastRow() - 1;
  const linhas = n > 0 ? sh.getRange(2, 1, n, CAB_CADASTROS.length).getDisplayValues() : [];
  const participantes = linhas.map(function (r) {
    return {
      data: r[0], nome: r[1], whatsapp: r[2], email: r[3], instagram: r[4],
      atuacao: r[5], cidade: r[6], segue: r[7], regulamento: r[8]
    };
  });
  return json({ ok: true, participantes: participantes });
}

function resultado(p) {
  if (p.senha !== SENHA) return json({ ok: false, erro: 'Senha incorreta.' });
  const sh = aba(ABA_SORTEIO, CAB_SORTEIO);
  sh.appendRow([agora(), texto(p.posicao), texto(p.nome), texto(p.whatsapp), texto(p.email),
    texto(p.instagram), texto(p.situacao), texto(p.consultora)]);
  return json({ ok: true });
}

/* ---------- utilitários ---------- */
function aba(nome, cab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh && nome === ABA_CADASTROS) {
    sh = ss.getSheets()[0];
    sh.setName(ABA_CADASTROS);
  }
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) sh.appendRow(cab);
  if (sh.getFrozenRows() === 0) {
    sh.getRange(1, 1, 1, cab.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function agora() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'); }
function limpa(v) { return String(v || '').trim().replace(/\s+/g, ' ').slice(0, 200); }
function soDigitos(v) { return String(v || '').replace(/\D/g, '').replace(/^55(?=\d{10,11}$)/, ''); }
function formataFone(d) { return '(' + d.slice(0, 2) + ') ' + d.slice(2, -4) + '-' + d.slice(-4); }
// impede que um texto seja interpretado como fórmula ou número pela planilha
function texto(v) { v = limpa(v); return /^[=+\-@(0-9]/.test(v) ? "'" + v : v; }
