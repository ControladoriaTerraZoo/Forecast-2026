const SPREADSHEET_ID = '1NuHqyLdpBjX7hfoNzDMHrmiaH9V6chj2agKXusmWX5w';
const SHEET_DRE   = 'DRE_BANCO';
const SHEET_REG   = 'Regiões';
const SHEET_USERS = 'USUARIOS';   // aba: nome | email | senha | papel | ativo

// Chave de acesso — precisa ser idêntica à do index.html (front-end).
// Fica visível a quem ler o código-fonte público; não é segurança forte,
// apenas evita uso casual da URL do backend por quem a encontrar solta.
const CHAVE_ACESSO = '7c19a4f2e8b6d035a1f94c72e0b8d6153ea4c9f7b2d81065';

// cod_conta únicos que compõem o forecast — evita dupla contagem.
// Corresponde 1:1 às 26 linhas do array LINHAS no index.html (mesma ordem).
const COD_FORECAST = [
  1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16,
  17, 18, 19, 20, 21, 22, 23, 24, 25, 26,
];

function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.chave !== CHAVE_ACESSO) {
      return _json({ status: 'erro', message: 'Acesso negado.' });
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // -- DRE (forecast) --
    const sheetDRE = ss.getSheetByName(SHEET_DRE);
    if (!sheetDRE) {
      return _json({ status: 'erro', message: 'Aba "' + SHEET_DRE + '" não encontrada na planilha.' });
    }
    const dataDRE = sheetDRE.getDataRange().getValues();
    const hDRE = dataDRE[0].map(function(x){ return String(x).trim(); });
    const iMes  = hDRE.indexOf('mes_ano');
    const iLoja = hDRE.indexOf('loja');
    const iCod  = hDRE.indexOf('cod_conta');
    const iDRE  = hDRE.indexOf('linha_dre');
    const iNiv  = hDRE.indexOf('nivel');
    const iReal = hDRE.indexOf('realizado');
    const iOrc  = hDRE.indexOf('orcado');

    const faltando = ['mes_ano', 'loja', 'cod_conta', 'linha_dre', 'realizado', 'orcado']
      .filter(function(nome){ return hDRE.indexOf(nome) < 0; });
    if (faltando.length) {
      return _json({ status: 'erro', message: 'Coluna(s) não encontrada(s) em "' + SHEET_DRE + '": ' + faltando.join(', ') });
    }

    const rows = dataDRE.slice(1)
      .filter(function(r){ return COD_FORECAST.indexOf(Number(r[iCod])) >= 0; })
      .map(function(r){ return {
        mes_ano:   _mesAno(r[iMes]),
        loja:      String(r[iLoja] || '').trim(),
        cod_conta: Number(r[iCod]  || 0),
        linha_dre: String(r[iDRE]  || '').trim(),
        nivel:     String(r[iNiv]  || 'grupo'),
        realizado: Number(r[iReal] || 0),
        orcado:    Number(r[iOrc]  || 0),
      };});

    // -- Regioes --
    const sheetREG = ss.getSheetByName(SHEET_REG);
    const regioes = {};
    if (sheetREG) {
      sheetREG.getDataRange().getValues().slice(1).forEach(function(r){
        if (r[0] && r[1]) regioes[String(r[0]).trim()] = String(r[1]).trim();
      });
    }

    // -- Usuarios (login validado no cliente) --
    const usuarios = [];
    const shU = ss.getSheetByName(SHEET_USERS);
    if (shU) {
      const dataU = shU.getDataRange().getValues();
      // Normaliza os cabeçalhos (minúsculo, sem acento/espaço/pontuação) e busca por
      // palavra-chave, pois a aba pode ter nomes como "E-mail (login)" ou "Senha Inicial"
      // em vez do nome exato da coluna.
      const hU = dataU[0].map(function(x){
        return String(x).trim().toLowerCase()
          .normalize('NFD').replace(/[̀-ͯ]/g, '')
          .replace(/[^a-z0-9]/g, '');
      });
      const _col = function(keywords){
        for (var i = 0; i < hU.length; i++) {
          for (var k = 0; k < keywords.length; k++) {
            if (hU[i].indexOf(keywords[k]) >= 0) return i;
          }
        }
        return -1;
      };
      const uNome  = _col(['nome']);
      const uEmail = _col(['email', 'mail']);
      const uSenha = _col(['senha']);
      const uPapel = _col(['papel', 'perfil', 'role']);
      const uAtivo = _col(['ativo', 'status']);
      dataU.slice(1).forEach(function(r){
        const email = uEmail >= 0 ? String(r[uEmail] || '').trim() : '';
        if (!email) return;
        usuarios.push({
          nome:  uNome  >= 0 ? String(r[uNome]  || '').trim() : email,
          email: email,
          senha: uSenha >= 0 ? String(r[uSenha] || '').trim() : '',
          papel: uPapel >= 0 ? String(r[uPapel] || 'leitor').trim().toLowerCase() : 'leitor',
          ativo: uAtivo >= 0 ? String(r[uAtivo] || 'sim').trim().toLowerCase() : 'sim',
        });
      });
    }

    return _json({ status: 'ok', rows: rows, regioes: regioes, usuarios: usuarios });

  } catch (err) {
    return _json({ status: 'erro', message: err.message });
  }
}

// Normaliza mes_ano para 'YYYY-MM', mesmo se a célula vier como Date
// (Sheets converte texto tipo "2026-01" para data automaticamente em alguns formatos).
function _mesAno(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(v || '').replace(/'/g, '').trim().slice(0, 7);
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
