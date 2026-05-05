/**
 * ============================================================
 * LOJA MARCIA — Sync Planilha → GitHub
 * ============================================================
 * Cole TODO este código no Apps Script da planilha.
 *
 * O que faz:
 *  1. Toda vez que você salvar uma edição na planilha,
 *     o site é atualizado automaticamente em ~1 minuto.
 *  2. Também sincroniza a cada 10 minutos (caso perca edição).
 *  3. Busca imagem no Bing quando você adiciona produto novo
 *     sem imagem na coluna F.
 * ============================================================
 */

// ---- CONFIGURAÇÃO ----
const GITHUB_TOKEN   = 'COLE_SEU_TOKEN_GITHUB_AQUI';
const GITHUB_REPO    = 'lojinhadd/loja-marcia';
const GITHUB_FILE    = 'data.csv';
const SHEET_NAME     = 'planilha_loja';
const SPREADSHEET_ID = '12X_rnVn7XUf5kjDliIlX6xPbceqZTHI7om0J04ZitOk';

// ---- COLUNAS ----
const COL_NOME      = 1; // A
const COL_PRECO     = 2; // B
const COL_ESTOQUE   = 3; // C
const COL_SALA      = 4; // D
const COL_CATEGORIA = 5; // E
const COL_IMAGEM    = 6; // F
const COL_ATIVO     = 7; // G

// ============================================================
// GATILHO DE EDIÇÃO — chamado automaticamente ao salvar
// ============================================================
function onEditInstalado(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < 2) return;

    // Se editou nome ou categoria sem imagem → busca imagem
    if (col === COL_NOME || col === COL_CATEGORIA) {
      const nome = sheet.getRange(row, COL_NOME).getValue();
      if (nome && String(nome).trim() !== '') {
        const imgAtual = sheet.getRange(row, COL_IMAGEM).getValue();
        if (!imgAtual || !String(imgAtual).startsWith('http')) {
          Utilities.sleep(500);
          const categoria = sheet.getRange(row, COL_CATEGORIA).getValue() || '';
          const imgUrl = buscarImagemBing(nome, categoria);
          if (imgUrl) {
            sheet.getRange(row, COL_IMAGEM).setValue(imgUrl);
          }
        }
      }
    }

    // Sincroniza com GitHub após qualquer edição
    sincronizarGitHub();

  } catch (err) {
    console.error('onEditInstalado erro:', err);
  }
}

// ============================================================
// SINCRONIZAÇÃO COM GITHUB
// ============================================================
function sincronizarGitHub() {
  try {
    garantirLinhaConfig();
    const csv = gerarCSV();
    const base64 = Utilities.base64Encode(Utilities.newBlob(csv, 'text/plain', 'data.csv').getBytes());

    const shaUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents/${GITHUB_FILE}`;
    const shaResp = UrlFetchApp.fetch(shaUrl, {
      headers: { 'Authorization': `token ${GITHUB_TOKEN}`, 'Accept': 'application/vnd.github.v3+json' },
      muteHttpExceptions: true
    });

    let sha = null;
    if (shaResp.getResponseCode() === 200) {
      sha = JSON.parse(shaResp.getContentText()).sha;
    }

    const body = {
      message: 'Auto-sync: planilha atualizada',
      content: base64,
      ...(sha && { sha })
    };

    const putResp = UrlFetchApp.fetch(shaUrl, {
      method: 'PUT',
      headers: {
        'Authorization': `token ${GITHUB_TOKEN}`,
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    const code = putResp.getResponseCode();
    if (code === 200 || code === 201) {
      console.log('✅ GitHub sincronizado com sucesso!');
      return true;
    } else {
      console.error('❌ Erro GitHub:', code, putResp.getContentText().substring(0, 200));
      return false;
    }

  } catch (err) {
    console.error('sincronizarGitHub erro:', err);
    return false;
  }
}

// ============================================================
// GERA O CSV A PARTIR DA PLANILHA
// ============================================================
function gerarCSV() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  const data  = sheet.getDataRange().getValues();

  const lines = data.map((row, i) => {
    while (row.length < 7) row.push('');

    return row.slice(0, 7).map((cell, j) => {
      let val;

      if (j === COL_PRECO - 1 || j === COL_ESTOQUE - 1) {
        if (cell instanceof Date) {
          val = '0';
        } else {
          val = String(cell ?? '').trim();
          if (j === COL_PRECO - 1 && val.endsWith('.0')) val = val.slice(0, -2);
        }
      } else if (cell instanceof Date) {
        val = Utilities.formatDate(cell, 'America/Sao_Paulo', 'yyyy-MM-dd');
      } else {
        val = String(cell ?? '').trim();
      }

      if (val.includes(',') || val.includes('"') || val.includes('\n')) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      return val;
    }).join(',');
  });

  return lines.join('\n');
}

// ============================================================
// BUSCA IMAGEM NO BING
// ============================================================
function buscarImagemBing(nome, categoria) {
  try {
    const query   = encodeURIComponent(nome + ' ' + (categoria || '') + ' produto');
    const url     = `https://www.bing.com/images/search?q=${query}&form=HDRSC2&first=1`;
    const headers = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0 Safari/537.36' };
    const resp    = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true });
    const html    = resp.getContentText();
    const match   = html.match(/murl&quot;:&quot;(https?:\/\/[^&"]+\.(jpg|jpeg|png|webp)[^&"]*)/i)
                 || html.match(/"murl":"(https?:\/\/[^"]+\.(jpg|jpeg|png|webp)[^"]*)"/i);
    if (match && match[1]) return match[1];
    return null;
  } catch (err) {
    console.error('Bing erro:', err);
    return null;
  }
}

// ============================================================
// BUSCAR IMAGENS EM LOTE
// ============================================================
function buscarImagensLote() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  const ultima = sheet.getLastRow();
  let count = 0;

  for (let row = 2; row <= ultima; row++) {
    const nome = sheet.getRange(row, COL_NOME).getValue();
    if (!nome || String(nome).trim() === '') continue;
    const img = sheet.getRange(row, COL_IMAGEM).getValue();
    if (img && String(img).startsWith('http')) continue;

    Utilities.sleep(400);
    const url = buscarImagemBing(String(nome), sheet.getRange(row, COL_CATEGORIA).getValue() || '');
    if (url) { sheet.getRange(row, COL_IMAGEM).setValue(url); count++; }
  }

  SpreadsheetApp.openById(SPREADSHEET_ID).toast(`✅ ${count} imagens preenchidas!`, 'Auto-imagem', 5);
  sincronizarGitHub();
}

// ============================================================
// LINHA DE CONFIG DO DESCONTO
// ============================================================
function garantirLinhaConfig() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_NOME - 1]).trim() === '__CONFIG__') {
      return false;
    }
  }

  sheet.insertRowBefore(2);
  const configRow = sheet.getRange(2, 1, 1, 8);
  configRow.setValues([[
    '__CONFIG__',
    25,
    '',
    '',
    'Desconto global do site (%)',
    '',
    'nao',
    ''
  ]]);

  configRow.setBackground('#fff9c4');
  sheet.getRange(2, 1).setFontWeight('bold');
  sheet.getRange(2, 2).setFontWeight('bold');
  sheet.getRange(2, 5).setFontColor('#888888');

  return true;
}

// ============================================================
// SINCRONIZAÇÃO MANUAL
// ============================================================
function sincronizarAgora() {
  const ok = sincronizarGitHub();
  SpreadsheetApp.openById(SPREADSHEET_ID).toast(
    ok ? '✅ Site atualizado com sucesso!' : '❌ Erro ao sincronizar. Verifique o log.',
    'Sincronizar', 5
  );
}

// ============================================================
// CONFIGURA TUDO — RODE UMA VEZ
// ============================================================
function configurarGatilhos() {
  // 1. Renomeia a aba para planilha_loja (se necessário)
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  if (sheets.length > 0 && sheets[0].getName() !== SHEET_NAME) {
    sheets[0].setName(SHEET_NAME);
  }

  // 2. Torna a planilha pública (leitura) para o site conseguir buscar os dados
  try {
    DriveApp.getFileById(SPREADSHEET_ID)
      .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    console.log('✅ Planilha tornada pública (somente leitura)');
  } catch (err) {
    console.error('Erro ao tornar planilha pública:', err);
  }

  // 3. Garante cabeçalhos na linha 1
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  const primeiraLinha = sheet.getRange(1, 1).getValue();
  if (!primeiraLinha || String(primeiraLinha).trim() === '') {
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Nome', 'Preco', 'Estoque', 'Sala', 'Categoria', 'Imagem', 'Ativo', 'Comprado por'
    ]]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#e11d48').setFontColor('#ffffff');
  }

  // 4. Garante linha __CONFIG__
  garantirLinhaConfig();

  // 5. Remove gatilhos antigos e recria
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEditInstalado')
    .forSpreadsheet(ss).onEdit().create();

  ScriptApp.newTrigger('sincronizarGitHub')
    .timeBased().everyMinutes(10).create();

  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(ss).onOpen().create();

  // 6. Sincroniza imediatamente
  sincronizarGitHub();

  ss.toast(
    '✅ Tudo configurado!\n\nO site será atualizado:\n• Toda vez que você editar\n• A cada 10 minutos automaticamente',
    'Setup completo', 10
  );
}

// ============================================================
// MENU PERSONALIZADO
// ============================================================
function onOpen() {
  SpreadsheetApp.openById(SPREADSHEET_ID)
    .addMenu('🛍️ Loja Marcia', [
      { name: '🔄 Sincronizar site agora',               functionName: 'sincronizarAgora'    },
      { name: '🔍 Buscar imagens em lote',               functionName: 'buscarImagensLote'   },
      { name: '⚙️  Reconfigurar gatilhos (1x)',          functionName: 'configurarGatilhos'  },
      { name: '⚙️  Criar linha de desconto na planilha', functionName: 'garantirLinhaConfig' }
    ]);

  garantirLinhaConfig();
}
