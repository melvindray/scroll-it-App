// Scroll-it · registrations + activity → this Google Sheet.
// Paste this whole file into Extensions > Apps Script, then Deploy > New deployment > Web app.
// To remove someone's access: in the "Inscrits" tab, write  non  in the "Accès" column of their row.

const TABS = {
  users: ['Inscrits', ['ID', 'Nom', 'Inscrit le', 'Vu pour la dernière fois', 'Accès', 'Appareil']],
  act: ['Activité', ['Date', 'ID', 'Nom', 'Film', 'Clip', 'Titre', 'Événement', 'Secondes']],
};

function sheet_(k) {
  const ss = SpreadsheetApp.getActive();
  const [name, head] = TABS[k];
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); sh.getRange(1, 1, 1, head.length).setFontWeight('bold'); }
  return sh;
}

function setup() {
  sheet_('users'); sheet_('act');
  const ss = SpreadsheetApp.getActive();
  let r = ss.getSheetByName('Résumé');
  if (!r) {
    r = ss.insertSheet('Résumé', 0);
    r.getRange('A1:B5').setValues([
      ['Inscrits', "=COUNTA(Inscrits!A2:A)"],
      ["Actifs aujourd'hui", "=IFERROR(COUNTUNIQUE(FILTER(Activité!B2:B, INT(Activité!A2:A)=TODAY())), 0)"],
      ['Actifs (7 jours)', "=IFERROR(COUNTUNIQUE(FILTER(Activité!B2:B, Activité!A2:A>=TODAY()-6)), 0)"],
      ['Minutes regardées (7 jours)', "=ROUND(SUMIFS(Activité!H2:H, Activité!A2:A, \">=\"&(TODAY()-6))/60)"],
      ['', ''],
    ]);
    r.getRange('A7').setValue('Clips les plus regardés');
    r.getRange('A8').setFormula("=IFERROR(QUERY(Activité!D2:H, \"select D, E, F, count(E), sum(H) where G = 'watch' group by D, E, F order by count(E) desc limit 20 label count(E) 'Vues', sum(H) 'Secondes'\", 0), \"\")");
    r.getRange('A1:A7').setFontWeight('bold');
    const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Feuille 1'); if (s1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  }
}

function findUser_(id) {
  const sh = sheet_('users'); const n = sh.getLastRow();
  if (n < 2) return null;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return { sh, row: i + 2 };
  return null;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  const id = (e.parameter.u || '').slice(0, 60);
  const u = id && findUser_(id);
  if (!u) return json_({ ok: true, known: false });
  const acc = String(u.sh.getRange(u.row, 5).getValue()).trim().toLowerCase();
  return json_({ ok: acc !== 'non', known: true });
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    setup();
    const d = JSON.parse(e.postData.contents || '{}');
    const id = String(d.u || '').slice(0, 60), name = String(d.name || '').slice(0, 80);
    if (!id) return json_({ ok: false });
    const now = new Date();
    let u = findUser_(id);
    if (!u) { sheet_('users').appendRow([id, name, now, now, 'oui', String(d.ua || '').slice(0, 60)]); u = findUser_(id); }
    else u.sh.getRange(u.row, 4).setValue(now);
    const ev = (d.ev || []).slice(0, 200).map(x => [new Date(x.t || now), id, name, String(x.f || ''), x.c || '', String(x.ti || '').slice(0, 80), String(x.e || ''), Math.round(+x.s || 0)]);
    if (ev.length) { const sh = sheet_('act'); sh.getRange(sh.getLastRow() + 1, 1, ev.length, 8).setValues(ev); }
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}
