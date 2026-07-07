/**
 * Posts Impulsionados — backend da Dashboard de Seguidores (Hire / Malu / Internacional)
 * Guarda os posts numa aba 'POSTS_DASH' DESTA planilha e serve pra dashboard.
 *
 * COMO PUBLICAR (faça 1 vez):
 * 1. Abra a planilha de SEGUIDORES (a que a dash já usa).
 * 2. Menu:  Extensões  →  Apps Script
 * 3. Apague tudo que estiver lá e COLE este arquivo inteiro.
 * 4. Clique no disquete (Salvar).
 * 5. Botão azul  Implantar  →  Nova implantação
 * 6. Na engrenagem ⚙ escolha:  App da Web
 * 7. Configurar:
 *      - Executar como:      Eu (seu email)
 *      - Quem tem acesso:    Qualquer pessoa
 * 8. Implantar → Autorizar acesso (aceite as permissões da sua conta).
 * 9. Vai aparecer uma "URL do app da Web" terminando em  /exec  → COPIE e mande pro Cardo.
 *
 * Pronto. O Cardo liga a dash nessa URL e os posts passam a sincronizar entre todos.
 */

const TAB = 'POSTS_DASH';

function _sheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if(!sh){
    sh = ss.insertSheet(TAB);
    sh.getRange(1,1,1,8).setValues([['profile','nome','link','inv','dia','saldo','status','obs']]);
  }
  return sh;
}

function _getPosts(profile){
  const rows = _sheet().getDataRange().getValues();
  const out = [];
  for(let i=1;i<rows.length;i++){
    const r = rows[i];
    if(!profile || String(r[0]) === profile){
      out.push({ nome:r[1], link:r[2], inv:Number(r[3])||0, dia:Number(r[4])||0, saldo:Number(r[5])||0, status:r[6]||'ativo', obs:r[7]||'' });
    }
  }
  return out;
}

// Leitura (a dash chama via JSONP: ...?profile=hire&callback=xxx)
function doGet(e){
  const p = (e.parameter.profile || '').trim();
  const payload = JSON.stringify({ ok:true, posts:_getPosts(p) });
  const cb = e.parameter.callback;
  if(cb) return ContentService.createTextOutput(cb + '(' + payload + ')')
                              .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

// Escrita (a dash envia {profile, posts:[...]} — substitui os posts daquele perfil)
function doPost(e){
  const body = JSON.parse(e.postData.contents);
  const profile = (body.profile || '').trim();
  const posts = body.posts || [];
  const sh = _sheet();
  const rows = sh.getDataRange().getValues();
  // remove os posts antigos deste perfil
  for(let i = rows.length - 1; i >= 1; i--){
    if(String(rows[i][0]) === profile) sh.deleteRow(i + 1);
  }
  // grava os novos
  const add = posts.map(p => [profile, p.nome||'', p.link||'', Number(p.inv)||0, Number(p.dia)||0, Number(p.saldo)||0, p.status||'ativo', p.obs||'']);
  if(add.length) sh.getRange(sh.getLastRow()+1, 1, add.length, 8).setValues(add);
  return ContentService.createTextOutput(JSON.stringify({ ok:true })).setMimeType(ContentService.MimeType.JSON);
}
