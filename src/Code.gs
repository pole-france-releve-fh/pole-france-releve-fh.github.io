/**
 * Pôle France Relève fleuret hommes · CREPS IDF · application web (v4).
 * Trois sources :
 *  - BASE (ex-SUIVI)  : ce qui se passe réellement (tireurs, séances, appel, scores, wellness, tests, matériel, budget, accès)
 *  - MAÎTRE           : l'organisation (semaines, créneaux, leçons, compétitions)
 *  - V6 (résultats)   : lecture seule, mise en cache chaque nuit dans Drive
 * Rôles (onglet Accès de BASE) : Staff, Tireur, Invité (résultats seulement).
 */

var IDS = {
  BASE: '1mi1W0UbrGYWRWTevamKS2C2_47Ugj3sKyW-Y0cJiyNc',
  MAITRE: '19OSdJs7yUOVvR9l1A5Il9XN2AToSSrrzp66IlWTAfC4',
  VACATIONS: '1nUTJPRhcbG7tYr62m0o1K6Z5MvPd8r3quKSpDkuEcJY',
  RESULTATS: '1sGQMWDfhNlt4HiD3rYI8PHwnp3bcM_ToUC28byVMG3w'
};
var TZ = 'Europe/Paris';
var CATS = ['M17 FH', 'M20 FH', 'M23 FH', 'Senior FH'];
var DOSSIER_APP = 'PFR • Fichiers de l\'app';

function doGet(e) {
  if (e && e.parameter && e.parameter.voix != null) return ContentService.createTextOutput(voixTexte_(e.parameter.k, e.parameter.voix)).setMimeType(ContentService.MimeType.TEXT);
  var t = HtmlService.createTemplateFromFile('Index');
  t.key = (e && e.parameter && e.parameter.k) || '';
  return t.evaluate()
    .setTitle('France Fleuret Hommes · Pôle France Relève')
    .setFaviconUrl('https://pole-france-releve-fh.github.io/favicon.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
 * Utilitaires
 * ======================================================= */

var _ss = {};
function ss_(id) { return _ss[id] || (_ss[id] = SpreadsheetApp.openById(id)); }
function sheet_(id, name) { return ss_(id).getSheetByName(name); }
function rows_(id, name) {
  var sh = sheet_(id, name);
  if (!sh || sh.getLastRow() < 1) return [];
  return sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), 1)).getDisplayValues();
}
function norm_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim();
}
function fmt_(d, f) { return Utilities.formatDate(d, TZ, f || 'dd/MM/yyyy'); }
function num_(x) { var n = Number(String(x).replace(/\s/g, '').replace(',', '.').replace(/[^\d.\-]/g, '')); return isNaN(n) ? 0 : n; }

/* Cache découpé (CacheService limite chaque valeur à 100 Ko). */
function cachePut_(k, obj, sec) {
  var s = JSON.stringify(obj), c = CacheService.getScriptCache(), n = Math.ceil(s.length / 90000), m = {};
  for (var i = 0; i < n; i++) m[k + '_' + i] = s.substr(i * 90000, 90000);
  m[k + '_n'] = String(n);
  try { c.putAll(m, sec || 300); } catch (e) { }
}
function cacheGet_(k) {
  var c = CacheService.getScriptCache(), n = Number(c.get(k + '_n'));
  if (!n) return null;
  var keys = []; for (var i = 0; i < n; i++) keys.push(k + '_' + i);
  var got = c.getAll(keys), s = '';
  for (var j = 0; j < n; j++) { if (got[keys[j]] == null) return null; s += got[keys[j]]; }
  try { return JSON.parse(s); } catch (e) { return null; }
}
function cacheClear_() { CacheService.getScriptCache().removeAll(['bundle_n', 'acces_n', 'cal_n', 'resr_n']); }

/* =========================================================
 * Accès
 * ======================================================= */

function acces_() {
  var a = cacheGet_('acces');
  if (a) return a;
  a = rows_(IDS.BASE, 'Accès').map(function (r, i) { return { nom: r[0], role: r[1], cle: r[2], email: String(r[3] || '').trim().toLowerCase(), rappel: r[4] === 'TRUE', sel: String(r[6] || ''), mdp: String(r[7] || ''), ligne: i + 1 }; })
    .slice(1).filter(function (x) { return x.cle; });
  cachePut_('acces', a, 600);
  return a;
}
function catsSel_(t) {
  return String(t || '').split(/[,;\s]+/).map(function (x) { x = x.trim(); if (!x) return ''; return (/^senior$/i.test(x) ? 'Senior' : x.toUpperCase()) + ' FH'; })
    .filter(function (x) { return CATS.indexOf(x) >= 0; });
}
function user_(key) {
  if (!key) return null;
  var u = acces_().filter(function (x) { return x.cle === key; })[0];
  if (!u) return null;
  var sc = catsSel_(u.sel), m = String(u.role || '').match(/^s[ée]lection(neur)?\s*(M17|M20|M23|Senior)?/i);
  if (m) { if (!sc.length && m[2]) sc = catsSel_(m[2]); return { nom: u.nom, role: 'Sélectionneur', officiel: !!m[1], selCats: sc, email: u.email }; }
  return { nom: u.nom, role: u.role || 'Invité', selCats: sc, email: u.email };
}
function staff_(key) { var u = user_(key); if (!u || u.role !== 'Staff') throw new Error('Réservé au staff'); return u; }
function selOk_(key, cat) {
  var u = user_(key);
  if (u && u.role === 'Staff') return u;
  if (u && u.role === 'Sélectionneur' && u.selCats.indexOf(cat) >= 0) return u;
  throw new Error('Accès non autorisé');
}
function panel_(cat) {
  return acces_().filter(function (a) {
    var sc = catsSel_(a.sel), m = String(a.role || '').match(/^s[ée]lection(neur)?\s*(M17|M20|M23|Senior)?/i);
    if (!sc.length && m && m[2]) sc = catsSel_(m[2]);
    return sc.indexOf(cat) >= 0;
  }).map(function (a) { return { nom: a.nom, officiel: /^s[ée]lectionneur/i.test(String(a.role || '')) }; });
}
/* Connexion par e-mail + mot de passe (empreinte salée SHA-256 en colonne H de l'onglet Accès) */
function hash_(salt, pwd) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pwd, Utilities.Charset.UTF_8)); }
function connexion(email, pwd) {
  email = String(email || '').trim().toLowerCase();
  var a = acces_().filter(function (x) { return x.email === email && x.mdp; })[0];
  if (a) { var p = a.mdp.split('$'); if (p.length === 2 && hash_(p[0], String(pwd || '')) === p[1]) return a.cle; }
  Utilities.sleep(1500);
  throw new Error('E-mail ou mot de passe incorrect');
}
function definirMotDePasse(key, email, pwd) {
  var u = acces_().filter(function (x) { return key && x.cle === key; })[0];
  if (!u) throw new Error('Lien invalide');
  email = String(email || '').trim().toLowerCase(); pwd = String(pwd || '');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Adresse e-mail invalide');
  if (pwd.length < 8) throw new Error('Le mot de passe doit faire au moins 8 caractères');
  if (acces_().some(function (x) { return x.email === email && x.cle !== key; })) throw new Error('Cette adresse est déjà utilisée par une autre personne');
  var salt = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  var sh = sheet_(IDS.BASE, 'Accès');
  sh.getRange(u.ligne, 4).setValue(email);
  sh.getRange(u.ligne, 8).setValue(salt + '$' + hash_(salt, pwd));
  cacheClear_();
  return 'ok';
}

/* =========================================================
 * Semaine (fichier MAÎTRE)
 * ======================================================= */

function hmin_(h) { var m = String(h).match(/(\d+)h?(\d*)/); return m ? Number(m[1]) * 60 + (Number(m[2]) || 0) : 9999; }
function isoWeek_(d) {
  var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  var day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  var y = t.getUTCFullYear();
  return { w: Math.ceil((((t - Date.UTC(y, 0, 1)) / 86400000) + 1) / 7), y: y };
}
function lundi_(offset) {
  var d = new Date(); d.setHours(12, 0, 0, 0);
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7) + 7 * (offset || 0));
  return d;
}
function weekTab_(d) {
  var k = isoWeek_(d), ss = ss_(IDS.MAITRE);
  return ss.getSheetByName('S' + (k.w < 10 ? '0' + k.w : k.w) + ' ' + k.y) || ss.getSheetByName('S' + k.w + ' ' + k.y);
}

function semaine_(offset) {
  var lun = lundi_(offset), sh = weekTab_(lun);
  if (!sh) return (offset || 0) < 0 && !gerees_()[fmt_(lun, 'yyyy-MM-dd')] ? null : semaineApp_(lun);
  var v = sh.getRange('A1:V90').getDisplayValues(), fc = sh.getRange('A1:V90').getFontColors();
  function vert_(r, c) { var m = String((fc[r] || [])[c] || '').match(/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i); if (!m) return false; var R = parseInt(m[1], 16), G = parseInt(m[2], 16), B = parseInt(m[3], 16); return G > 120 && G - R > 60 && G - B > 60; }
  var dates = [];
  for (var i = 0; i < 7; i++) { var d = new Date(lun); d.setDate(lun.getDate() + i); dates.push(fmt_(d, 'yyyy-MM-dd')); }

  // séances collectives : trois blocs (matin, après-midi, soir)
  var jours = [];
  for (var c = 1, j = 0; c <= 21; c += 4, j++) {
    var titre = v[8][c];
    if (!titre) continue;
    var cren = [];
    [[12, 13, 14, 15, 'm'], [24, 25, 26, 27, 'a'], [29, 30, 31, 32, 's']].forEach(function (b) {
      if (v[b[2]][c] && v[b[2]][c] !== 'REPOS') cren.push({ h: v[b[0]][c], lieu: v[b[1]][c], quoi: v[b[2]][c], note: v[b[3]] ? v[b[3]][c] : '', slot: b[4] });
    });
    // créneaux INSEP : lieu en colonne B du jour, heure juste à droite, activité dessous
    [19, 36].forEach(function (r0) {
      var q = v[r0 + 1] ? v[r0 + 1][c] : '';
      if (q && q !== '-' && q !== 'REPOS') cren.push({ h: v[r0][c + 1] || '', lieu: v[r0][c] || 'INSEP', quoi: q, note: v[r0 + 2] ? v[r0 + 2][c] : '', slot: r0 === 19 ? 'i1' : 'i2' });
    });
    cren.sort(function (a, b) { return hmin_(a.h) - hmin_(b.h); });
    jours.push({ jour: titre, date: dates[j], evt: v[9][c], creneaux: cren, lecons: [], c: c });
  }

  // leçons individuelles : repérées par les libellés de la colonne B
  var iL = -1, iM = -1, iAM = -1, iNb = -1;
  for (var r = 30; r < v.length; r++) {
    var b = String(v[r][1]).trim().toUpperCase();
    if (b === 'LEÇONS INDIVIDUELLES') iL = r;
    else if (iL > 0 && iM < 0 && b.indexOf('MATIN') === 0) iM = r;
    else if (iL > 0 && b.indexOf('APRÈS MIDI') === 0) iAM = r;
    else if (iL > 0 && b.indexOf('NB LEÇONS') === 0) { iNb = r; break; }
  }
  if (iL > 0 && iM > 0) {
    var blocs = [{ coach: iM + 1, de: iM + 2, a: (iAM > 0 ? iAM : (iNb > 0 ? iNb : v.length)) - 1, moment: 'matin' }];
    if (iAM > 0) blocs.push({ coach: iAM + 1, de: iAM + 2, a: (iNb > 0 ? iNb : v.length) - 1, moment: 'après-midi' });
    blocs.forEach(function (bl) {
      for (var c2 = 1; c2 <= 21; c2 += 4) {
        var J = jours.filter(function (x) { return x.c === c2; })[0];
        if (!J) continue;
        for (var k = 0; k < 4 && c2 + k < 22; k++) {
          var coach = v[bl.coach] ? v[bl.coach][c2 + k] : '';
          var last = null;
          for (var rr = bl.de; rr < bl.a; rr += 2) {
            var h = v[rr] ? v[rr][c2 + k] : '', n = v[rr + 1] ? v[rr + 1][c2 + k] : '';
            if (!n) continue;
            if (/^\s*VS\s*$/i.test(h) && last) { last.vs = n; last.doublette = true; continue; }
            last = { moment: bl.moment, coach: coach, h: h, tireur: n };
            if (vert_(rr + 1, c2 + k)) { last.entretien = true; (J.entretiens = J.entretiens || []).push(last); last = null; continue; }
            J.lecons.push(last);
          }
        }
      }
    });
  }
  // dimanche : absent du fichier maître, reconstruit à partir des lignes « (D) » du samedi
  var sam = jours.filter(function (x) { return x.c === 21; })[0];
  var dd = new Date(lun); dd.setDate(lun.getDate() + 6);
  var MOISU = ['JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE'];
  var dim = { jour: 'DIMANCHE ' + dd.getDate() + ' ' + MOISU[dd.getMonth()], date: dates[6], evt: '', creneaux: [], lecons: [], c: 25, dimanche: true };
  if (sam && sam.evt) {
    var ls = String(sam.evt).split('\n'), garde = [], dl = [];
    ls.forEach(function (l) { if (/\(D\)\s*$/.test(l)) dl.push(l.replace(/\s*\(D\)\s*$/, '')); else garde.push(l); });
    sam.evt = garde.join('\n'); dim.evt = dl.join('\n');
  }
  jours.push(dim);
  var compets = [];
  for (var r2 = 3; r2 < 8; r2++) if (v[r2][14]) compets.push({ date: v[r2][13], nom: v[r2][14], lieu: v[r2][16], depart: v[r2][18], retour: v[r2][19] });
  var suivantes = [];
  for (var r3 = 3; r3 < 8; r3++) if (v[r3][21]) suivantes.push(v[r3][21]);
  // semaine gérée dans l'application : créneaux et leçons viennent de BASE (le fichier maître n'en est qu'une copie)
  var ge = !!gerees_()[dates[0]];
  if (ge) {
    var bc = crenBase_(dates[0]), lpv = leconsPlanifiees_(dates[0]).filter(function (x) { return /valid/i.test(x.statut); });
    jours.forEach(function (J) {
      if (J.dimanche) { J.lecons = []; J.entretiens = []; return; }
      J.creneaux = []; SLOTS_.forEach(function (sl) { var c = bc[J.date + '|' + sl]; if (c) J.creneaux.push({ h: c.h, lieu: c.lieu, quoi: c.quoi, note: c.note, slot: sl }); });
      J.creneaux.sort(function (a, b) { return hmin_(a.h) - hmin_(b.h); });
      var L = lpv.filter(function (x) { return x.date === J.date; });
      J.lecons = L.filter(function (x) { return !/entretien/i.test(x.note); }).map(lecApp_);
      J.entretiens = L.filter(function (x) { return /entretien/i.test(x.note); }).map(lecApp_);
    });
  }
  // leçons planifiées dans le portail (validées) : ajoutées à celles du fichier maître (sans doublon)
  if (!ge) try {
    var lp = leconsPlanifiees_(dates[0]).filter(function (x) { return /valid/i.test(x.statut); });
    var cle_ = function (d, t, h) { return d + '|' + norm_(t).split(' ')[0] + '|' + hmin_(h); };
    jours.forEach(function (J) {
      J.entretiens = J.entretiens || [];
      var vus = {};
      J.lecons.concat(J.entretiens).forEach(function (x) { vus[cle_(J.date, x.tireur, x.h)] = 1; });
      lp.filter(function (x) { return x.date === J.date; }).forEach(function (x) {
        if (vus[cle_(J.date, x.tireur, x.h)]) return;
        J.lecons.push({ moment: hmin_(x.h) < 13 * 60 ? 'matin' : 'après-midi', coach: x.coach, h: x.h, tireur: x.tireur, vs: x.vs, doublette: !!x.vs });
      });
    });
  } catch (e) { }
  return appliqueComp_({
    onglet: sh.getName(), gere: ge, titre: v[0][1], lundi: dates[0],
    assauts: v[3][4], pp: v[3][7], fonda: v[3][11], leconsPol: v[4][4], leconsPe: v[4][7],
    jours: jours, compets: compets, suivantes: suivantes
  }, dates);
}

/* =========================================================
 * Blocs de BASE
 * ======================================================= */

function vacations_() {
  var r = rows_(IDS.VACATIONS, 'Synthèse'), out = [], aValider = '';
  r.forEach(function (x) {
    if (['FFE', 'LIFE', 'TOTAL'].indexOf(x[0]) >= 0 && x[1]) out.push({ fin: x[0], env: x[1], conso: x[2], reste: x[3], taux: x[4], euros: x[5] });
    if (x[3] === 'Lignes à compléter / valider') aValider = x[5];
  });
  return { lignes: out, aValider: aValider };
}

function budget_() {
  var r = rows_(IDS.BASE, 'Budget commission'), dep = [];
  for (var i = 7; i < r.length; i++) {
    if (!r[i][0]) continue;
    dep.push({ code: r[i][0], nom: r[i][1], type: r[i][2], cat: r[i][3], date: r[i][4], statut: r[i][5], budget: r[i][6], att: r[i][7], ecart: r[i][8] });
  }
  var raw = sheet_(IDS.BASE, 'Budget commission').getRange('B3:H4').getValues();
  var pole = rows_(IDS.BASE, 'Compte pôle');
  return {
    enveloppe: raw[0][0], atterrissage: raw[0][3], marge: raw[0][6], solde: raw[1][0], reste: raw[1][3], pct: raw[1][6],
    deplacements: dep,
    pole: { solde: pole[1] ? pole[1][1] : '', lamesDues: pole[1] ? pole[1][4] : '', mouvements: pole.slice(4).filter(function (x) { return x[0] || x[1]; }).slice(-10).reverse() }
  };
}

function materiel_() {
  var p = sheet_(IDS.BASE, 'Paramètres').getRange('A4:B18').getValues(), par = {};
  p.forEach(function (x) { par[x[0]] = x[1]; });
  var lames = rows_(IDS.BASE, 'Lames').slice(1).filter(function (x) { return x[0]; })
    .map(function (x) { return { n: x[0], couleur: x[1], origine: x[2], prix: x[3], statut: x[4], tireur: x[5], date: x[6], paye: x[7] === 'TRUE', reste: x[8] }; });
  var autres = rows_(IDS.BASE, 'Matériel').slice(1).filter(function (x) { return x[0]; })
    .map(function (x) { return { article: x[0], stock: x[1], seuil: x[2], commander: x[3], comm: x[4] }; });
  return {
    total: par['Lames au total'], ffe: par['Dont dotation FFE'], pole: par['Dont achat pôle'], stock: par['Lames en stock'],
    vendues: par['Lames vendues'], aEncaisser: par['Montant HT restant à encaisser (€)'], alerte: par['Alerte stock'],
    seuil: par["Seuil d'alerte lames en stock"], prixBleue: par['Prix lame bleue HT (€)'], prixBlanche: par['Prix lame blanche HT (€)'],
    lames: lames, autres: autres
  };
}

var CHAMPS_FICHE = ['nom', 'prenom', 'annee', 'cat', 'club', 'ligue', 'statut', 'niveau', 'etudes', 'regime', 'email', 'main', 'taille', 'poids', 'ce',
  'naissance', 'envergure', 'morpho', 'tshirt', 'sweat', 'veste', 'pointure', 'photo', 'tel', 'parent', 'telParent'];
var LIB_FICHE = { naissance: 'Date de naissance', main: 'Main (latéralité)', taille: 'Taille (cm)', poids: 'Poids (kg)', envergure: 'Envergure (cm)', morpho: 'Morphologie / notes',
  tshirt: 'Taille t-shirt', sweat: 'Taille sweat / survêt', veste: 'Taille veste escrime', pointure: 'Pointure', tel: 'Téléphone', parent: 'Contact parent', telParent: 'Téléphone parent', email: 'E-mail' };

function tireurs_() {
  return rows_(IDS.BASE, 'Tireurs').slice(1).filter(function (x) { return x[0]; }).map(function (x, i) {
    var t = { ligne: i + 2 };
    CHAMPS_FICHE.forEach(function (c, j) { t[c] = x[j] || ''; });
    t.cle = norm_(t.nom + ' ' + t.prenom);
    t.pole = String(t.statut).indexOf('Pôle') === 0;
    t.peInsep = /^(oui|x|1|true|pe)/i.test(String(x[26] || '').trim());
    return t;
  });
}
function trouveTireur_(nom) {
  var k = norm_(nom);
  return tireurs_().filter(function (t) { return t.cle === k || norm_(t.prenom + ' ' + t.nom) === k || norm_(t.nom) === k; })[0] || null;
}

function wellness_(jours) {
  var sh = sheet_(IDS.BASE, 'Wellness');
  var raw = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues() : [];
  var lim = new Date(Date.now() - (jours || 14) * 86400000);
  return raw.filter(function (x) { return x[0] instanceof Date && x[0] >= lim; }).map(function (x) {
    return { d: fmt_(x[0], 'yyyy-MM-dd'), tireur: x[1], s: x[2], f: x[3], c: x[4], hu: x[5], st: x[6], total: x[7], comm: x[8] };
  });
}

function assautsBruts_() {
  var sh = sheet_(IDS.BASE, 'Assauts');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues().filter(function (x) { return x[2] && x[3]; }).map(function (x) {
    return { date: x[0] instanceof Date ? fmt_(x[0], 'yyyy-MM-dd') : String(x[0]), seance: x[1], a: x[2], b: x[3], sa: Number(x[4]) || 0, sb: Number(x[5]) || 0, format: x[7], type: x[9], idS: x[10], idE: x[11] };
  });
}
function assauts_() {
  var par = {}, list = assautsBruts_();
  function add(n, td, tr) {
    var k = norm_(n); par[k] = par[k] || { m: 0, v: 0, d: 0, td: 0, tr: 0 };
    par[k].m++; par[k].td += td; par[k].tr += tr; if (td > tr) par[k].v++; else if (td < tr) par[k].d++;
  }
  list.forEach(function (x) { add(x.a, x.sa, x.sb); add(x.b, x.sb, x.sa); });
  return { parTireur: par, derniers: list.slice(-20).reverse() };
}

function scolaire_() {
  return rows_(IDS.BASE, 'Scolaire').slice(1).filter(function (x) { return x[0]; })
    .map(function (x) { return { tireur: x[0], date: x[1], matiere: x[2], type: x[3], besoin: x[4], statut: x[5] }; });
}

function tests_() {
  return rows_(IDS.BASE, 'Tests physiques').slice(1).filter(function (x) { return x[1] && x[3]; })
    .map(function (x) { return { date: x[0], tireur: x[1], famille: x[2], test: x[3], val: num_(x[4]), unite: x[5], cote: x[6], haut: String(x[7]).toLowerCase() !== 'non', eval: x[8], rq: x[9], t: dateNum_(x[0]) }; })
    .sort(function (p, q) { return p.t - q.t; });
}
function dettes_() {
  return rows_(IDS.BASE, 'Dettes tireurs').slice(1).filter(function (x) { return x[1]; })
    .map(function (x, i) { return { ligne: i + 2, date: x[0], tireur: x[1], lib: x[2], montant: num_(x[3]), paye: x[4] === 'TRUE' || /^oui$/i.test(x[4]), datePaye: x[5], moyen: x[6], note: x[7] }; });
}
function demandes_() {
  return rows_(IDS.BASE, 'Demandes fiche').slice(1).map(function (x, i) { return { ligne: i + 2, date: x[0], tireur: x[1], champ: x[2], avant: x[3], apres: x[4], statut: x[5] }; })
    .filter(function (x) { return x.tireur && x.statut === 'À valider'; });
}

/* Résumé des classements par tireur (onglets « Vue » de V6, légers). */
function resultatsResume_() {
  var c = cacheGet_('resr');
  if (c) return c;
  var out = {};
  function evo(x) { var t = String(x || '').replace(/\u2212/g, '-'); if (/ERROR/.test(t)) return null; var m = t.match(/([+\-]?)\s?(\d+)/); return m ? (m[1] === '-' ? -1 : 1) * Number(m[2]) : null; }
  function put(nom, prenom, cat, o) {
    var k1 = norm_(nom + ' ' + prenom), k2 = norm_(String(nom).split(/[\s-]+/)[0] + ' ' + prenom);
    [k1, k2].forEach(function (k) { var x = out[k] = out[k] || {}; if (!x[cat] || !x[cat].clt) x[cat] = o; });
    var k4 = 'NI ' + norm_(nom) + ' ' + norm_(prenom).charAt(0); var z = out[k4] = out[k4] || {}; if (!z[cat] || !z[cat].clt) z[cat] = o;
    var k3 = 'NOM ' + norm_(nom); if (out[k3] === undefined || (out[k3] && out[k3]._p === norm_(prenom))) { var y = out[k3] = out[k3] || { _p: norm_(prenom) }; if (!y[cat] || !y[cat].clt) y[cat] = o; } else out[k3] = null;
  }
  var REF = 'depuis la dernière mise à jour';
  function lire(nomOnglet, n) { var sh = sheet_(IDS.RESULTATS, nomOnglet), l = sh.getLastRow(); return l < 5 ? [] : sh.getRange(5, 1, l - 4, n).getValues(); }
  function o(r, i) { return { clt: r[i], sansER: r[i + 1], evol: evo(r[i + 2]), ref: REF }; }
  lire('Senior FH', 10).forEach(function (r) { if (r[2]) put(r[2], r[3], 'Senior', o(r, 7)); });
  lire('M20 FH', 13).forEach(function (r) { if (!r[2]) return; put(r[2], r[3], 'M20', o(r, 7)); if (r[10]) put(r[2], r[3], 'Senior', o(r, 10)); });
  lire('M17 FH', 16).forEach(function (r) { if (!r[2]) return; put(r[2], r[3], 'M17', o(r, 7)); if (r[10]) put(r[2], r[3], 'M20', o(r, 10)); if (r[13]) put(r[2], r[3], 'Senior', o(r, 13)); });
  cachePut_('resr', out, 3600);
  return out;
}

/* =========================================================
 * Séances, appel, exercices (BASE)
 * ======================================================= */

function seancesSemaine_(lundiIso) {
  var fin = new Date(lundiIso + 'T12:00:00'); fin.setDate(fin.getDate() + 7);
  var finIso = fmt_(fin, 'yyyy-MM-dd');
  function iso(x) { return x instanceof Date ? fmt_(x, 'yyyy-MM-dd') : String(x); }
  var s = sheet_(IDS.BASE, 'Séances'), a = sheet_(IDS.BASE, 'Appel'), e = sheet_(IDS.BASE, 'Exercices');
  var se = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 14).getValues() : [];
  var ap = a.getLastRow() > 1 ? a.getRange(2, 1, a.getLastRow() - 1, 9).getValues() : [];
  var ex = e.getLastRow() > 1 ? e.getRange(2, 1, e.getLastRow() - 1, 12).getValues() : [];
  var list = se.filter(function (x) { var d = iso(x[1]); return x[0] && d >= lundiIso && d < finIso; }).map(function (x) {
    return { id: String(x[0]), date: iso(x[1]), h: String(x[2]), duree: x[3], type: x[4], lieu: x[5], coach: x[6], consignes: x[7], intensite: x[8], repos: x[9], photo: x[10], notes: x[11] };
  });
  var ids = {}; list.forEach(function (x) { ids[x.id] = x; x.appel = []; x.exercices = []; });
  ap.forEach(function (x) { var id = String(x[0]); if (ids[id]) ids[id].appel.push({ tireur: x[2], statut: x[3], retard: x[4], duree: x[5], motif: x[6] }); });
  ex.forEach(function (x) { var id = String(x[0]); if (ids[id]) ids[id].exercices.push({ idE: String(x[1]), ordre: x[2], format: x[3], touches: x[4], nb: x[5], intensite: x[6], repos: x[7], consignes: x[8], participants: x[9], n5: x[10], n15: x[11] }); });
  var matchs = assautsBruts_().filter(function (m) { return m.idS && ids[m.idS]; });
  matchs.forEach(function (m) { (ids[m.idS].matchs = ids[m.idS].matchs || []).push(m); });
  return list;
}

/* ---- Bilan de la saison : toutes les semaines depuis la rentrée (onglets semaine du fichier maître + leçons validées du portail) ---- */
var DEBUT_SAISON = '2026-08-31';
function getSaison(key, forcer) {
  staff_(key);
  var c = forcer ? null : cacheGet_('saison'); if (c) return c;
  var d0 = new Date(DEBUT_SAISON + 'T12:00:00'), off0 = Math.round((d0 - lundi_(0)) / (7 * 864e5)), sem = [];
  for (var o = off0; o <= 0; o++) {
    var s = null; try { s = semaine_(o); } catch (e) { }
    if (!s) continue;
    var pc = {}, pt = {}, nl = 0, nd = 0, ne = 0, ncoll = 0;
    s.jours.forEach(function (j) {
      j.creneaux.forEach(function (c) { if (!/^LE[CÇ]ONS?$/i.test(String(c.quoi).trim())) ncoll++; });
      j.lecons.forEach(function (l) {
        var co = String(l.coach || '?').trim(); pc[co] = (pc[co] || 0) + 1;
        [l.tireur, l.vs].forEach(function (t) { t = String(t || '').trim(); if (t) pt[t] = (pt[t] || 0) + 1; });
        if (l.vs) nd++; else nl++;
      });
      ne += (j.entretiens || []).length;
    });
    sem.push({ titre: s.titre, lundi: s.lundi, onglet: s.onglet, assauts: num_(s.assauts), pp: num_(s.pp), fonda: num_(s.fonda), collectifs: ncoll, lecons: nl, doublettes: nd, entretiens: ne, parCoach: pc, parTireur: pt, compets: s.compets.map(function (c) { return c.nom; }) });
  }
  var out = { semaines: sem, profils: profilsLecons_().map(function (p) { return { nom: p.nom, prenom: p.prenom, role: p.role, groupe: p.groupe, cible: p.cible }; }), maj: fmt_(new Date(), 'dd/MM/yyyy HH:mm') };
  cachePut_('saison', out, 6 * 3600);
  return out;
}
/* ---- Séances collectives : ajout, modification, suppression directement dans l'onglet de la semaine du fichier maître ----
 * Par jour (4 colonnes, lundi = colonne B) : matin (lignes 13-16), après-midi (25-28), soir (30-33) : heure / salle / activité / public ;
 * créneaux INSEP (lignes 20-22 et 37-39) : salle + heure à droite / activité / public. */
var CRENEAU_ROWS = { m: 12, a: 24, s: 29, i1: 19, i2: 36 };
function saveCreneau(key, lundi, date, slot, f) {
  var u = staff_(key);
  if (CRENEAU_ROWS[slot] == null) throw new Error('Créneau inconnu');
  var lun = new Date(lundi + 'T12:00:00'), j = Math.round((new Date(date + 'T12:00:00') - lun) / 864e5);
  if (j < 0 || j > 5) throw new Error('Seuls les jours du lundi au samedi se modifient ici');
  var lock = LockService.getScriptLock(), r = null; lock.waitLock(30000);
  try {
    basculer_(lundi, u.nom);
    crenUpsert_(lundi, date, slot, f, u.nom);
    r = ecrireMaitre_(lundi);
  } finally { lock.releaseLock(); }
  cacheClear_(); CacheService.getScriptCache().remove('saison_n');
  var s = semaine_(Math.round((lun - lundi_(0)) / (7 * 864e5)));
  if (s && r && r.horsGrille && r.horsGrille.length) s.horsGrille = r.horsGrille;
  return s;
}
/* =========================================================
 * Planning de la semaine géré dans l'application (étape 1 du détachement du fichier maître)
 * BASE › Créneaux : séances collectives par semaine · Semaine type : modèle · Semaines gérées : semaines basculées.
 * Pendant la transition, chaque modification est recopiée dans l'onglet « Sxx » du fichier maître (créneaux + grille
 * des leçons validées) ; une modification faite à la main dans le fichier maître est reprise dans l'application.
 * ======================================================= */
var SLOTS_ = ['m', 'i1', 'a', 's', 'i2'];
var JOURS_ = ['LUNDI', 'MARDI', 'MERCREDI', 'JEUDI', 'VENDREDI', 'SAMEDI', 'DIMANCHE'];
var MOIS_U = ['JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE'];
function shCree_(nom, ent) {
  var ss = ss_(IDS.BASE), sh = ss.getSheetByName(nom);
  if (!sh) { sh = ss.insertSheet(nom); sh.getRange(1, 1, 1, ent.length).setValues([ent]).setFontWeight('bold'); sh.setFrozenRows(1); sh.getRange('A:C').setNumberFormat('@'); }
  return sh;
}
function crenSh_() { return shCree_('Créneaux', ['Semaine (lundi)', 'Date', 'Créneau', 'Heure', 'Salle', 'Activité', 'Public', 'Mis à jour', 'Par']); }
function typeSh_() { return shCree_('Semaine type', ['Jour (0 = lundi)', 'Créneau', 'Heure', 'Salle', 'Activité', 'Public']); }
function gereSh_() { return shCree_('Semaines gérées', ['Semaine (lundi)', 'Gérée depuis', 'Par']); }
var _gerees = null;
function gerees_() {
  if (_gerees) return _gerees;
  var sh = gereSh_(), o = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().forEach(function (x) { if (x[0]) o[isoTxt_(x[0])] = 1; });
  return (_gerees = o);
}
function datesSem_(lundi) { var lun = new Date(lundi + 'T12:00:00'), o = []; for (var i = 0; i < 7; i++) { var d = new Date(lun); d.setDate(lun.getDate() + i); o.push(fmt_(d, 'yyyy-MM-dd')); } return o; }
function vide_(q) { q = String(q || '').trim(); return !q || q === '-' || q.toUpperCase() === 'REPOS'; }
function sigC_(c) { return c && !vide_(c.quoi) ? [c.h, c.lieu, c.quoi, c.note].map(function (x) { return String(x || '').trim(); }).join('|') : ''; }
function crenBase_(lundi) {
  var sh = crenSh_(), o = {};
  if (sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getDisplayValues().forEach(function (x, i) {
    if (isoTxt_(x[0]) !== lundi || vide_(x[5])) return;
    var d = isoTxt_(x[1]); o[d + '|' + x[2]] = { date: d, slot: x[2], h: x[3], lieu: x[4], quoi: x[5], note: x[6], ligne: i + 2 };
  });
  return o;
}
function crenType_() {
  var sh = typeSh_(), o = [];
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues().forEach(function (x) { if (x[1] && !vide_(x[4])) o.push({ j: Number(x[0]), slot: x[1], h: x[2], lieu: x[3], quoi: x[4], note: x[5] }); });
  return o;
}
/* Remplace tous les créneaux d'une semaine dans BASE (liste de {date, slot, h, lieu, quoi, note}). */
function crenRemplace_(lundi, liste, par) {
  var sh = crenSh_();
  if (sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
    for (var i = v.length - 1; i >= 0; i--) if (isoTxt_(v[i][0]) === lundi) sh.deleteRow(i + 2);
  }
  var rows = liste.filter(function (c) { return !vide_(c.quoi); }).map(function (c) { return [lundi, c.date, c.slot, c.h || '', c.lieu || '', c.slot.charAt(0) === 'i' ? c.quoi : String(c.quoi).toUpperCase(), c.note || '', new Date(), par || '']; });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
}
function crenUpsert_(lundi, date, slot, f, par) {
  var sh = crenSh_(), b = crenBase_(lundi), ex = b[date + '|' + slot];
  if (!f || vide_(f.quoi)) { if (ex) sh.deleteRow(ex.ligne); return; }
  var row = [lundi, date, slot, String(f.h || '').trim(), String(f.lieu || '').trim(), String(f.quoi).trim().toUpperCase(), String(f.note || '').trim(), new Date(), par || ''];
  if (ex) sh.getRange(ex.ligne, 1, 1, 9).setValues([row]); else sh.appendRow(row);
}

/* ---- lecture brute d'un onglet « Sxx » du fichier maître ---- */
function maitreBrut_(sh) {
  var rg = sh.getRange('A1:Y90'), v = rg.getDisplayValues(), fc = rg.getFontColors();
  var L = { iL: -1, iM: -1, iAM: -1, iNb: -1 };
  for (var r = 30; r < v.length; r++) {
    var b = String(v[r][1]).trim().toUpperCase();
    if (b === 'LEÇONS INDIVIDUELLES') L.iL = r;
    else if (L.iL > 0 && L.iM < 0 && b.indexOf('MATIN') === 0) L.iM = r;
    else if (L.iL > 0 && b.indexOf('APRÈS MIDI') === 0) L.iAM = r;
    else if (L.iL > 0 && b.indexOf('NB LEÇONS') === 0) { L.iNb = r; break; }
  }
  return { v: v, fc: fc, L: L };
}
function vertC_(col) { var m = String(col || '').match(/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i); if (!m) return false; var R = parseInt(m[1], 16), G = parseInt(m[2], 16), B = parseInt(m[3], 16); return G > 120 && G - R > 60 && G - B > 60; }
function maitreCren_(v, dates) {
  var o = {};
  for (var j = 0; j < 6; j++) {
    var c = 1 + 4 * j;
    [['m', 12], ['a', 24], ['s', 29]].forEach(function (b) { var r = b[1]; o[dates[j] + '|' + b[0]] = { h: v[r][c], lieu: v[r + 1][c], quoi: v[r + 2][c], note: v[r + 3][c] }; });
    [['i1', 19], ['i2', 36]].forEach(function (b) { var r = b[1], l = v[r][c]; o[dates[j] + '|' + b[0]] = { h: v[r][c + 1], lieu: l === 'INSEP' ? 'INSEP' : l, quoi: v[r + 1][c], note: v[r + 2][c] }; });
  }
  return o;
}
function blocs_(L) {
  var o = [];
  if (L.iM > 0 && L.iAM > 0) o.push({ moment: 'matin', coach: L.iM + 1, de: L.iM + 2, a: L.iAM - 1 });
  if (L.iAM > 0 && L.iNb > 0) o.push({ moment: 'après-midi', coach: L.iAM + 1, de: L.iAM + 2, a: L.iNb - 1 });
  o.forEach(function (b) { b.n = Math.ceil((b.a - b.de) / 2); });
  return o;
}
function cc_(x) { return norm_(x).replace(/ /g, '').slice(0, 4); }
function cleL_(date, coach, h, t) { return date + '|' + cc_(coach) + '|' + hmin_(h) + '|' + norm_(t); }
function maitreLecons_(B, dates) {
  var out = [], v = B.v;
  blocs_(B.L).forEach(function (bl) {
    for (var j = 0; j < 6; j++) for (var k = 0; k < 4; k++) {
      var col = 1 + 4 * j + k, coach = v[bl.coach][col], last = null;
      for (var i = 0; i < bl.n; i++) {
        var rr = bl.de + 2 * i, h = v[rr][col], n = String(v[rr + 1][col] || '').trim();
        if (!n) continue;
        if (/^\s*VS\s*$/i.test(h) && last) { last.vs = n; continue; }
        last = { date: dates[j], moment: bl.moment, coach: coach, h: h, tireur: n, vs: '', entretien: vertC_(B.fc[rr + 1][col]) };
        out.push(last);
      }
    }
  });
  return out;
}
/* Le fichier maître peut imposer une liste de valeurs pour la case « public » : sinon elle reste vide (l'application garde le détail). */
function noteMaitre_(sh, j, slot, note) {
  note = String(note || '').trim(); if (!note) return '';
  var r0 = CRENEAU_ROWS[slot], row = slot.charAt(0) === 'i' ? r0 + 3 : r0 + 4;
  try {
    var dv = sh.getRange(row, 2 + 4 * j).getDataValidation();
    if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) { var L = dv.getCriteriaValues()[0] || []; return L.indexOf(note) >= 0 ? note : ''; }
  } catch (e) { }
  return note;
}
function ecrireSlot_(sh, j, slot, f) {
  var r0 = CRENEAU_ROWS[slot], col = 2 + 4 * j, t = function (x) { return String(x == null ? '' : x).trim(); };
  if (slot.charAt(0) === 'i') {
    sh.getRange(r0 + 1, col, 1, 2).setValues([[f ? (t(f.lieu) || 'INSEP') : 'INSEP', f ? t(f.h) : '']]);
    sh.getRange(r0 + 2, col, 2, 1).setValues([[f ? (t(f.quoi) || '-') : '-'], [f ? t(f.note) : '']]);
  } else sh.getRange(r0 + 1, col, 4, 1).setValues(f ? [[t(f.h)], [t(f.lieu)], [t(f.quoi).toUpperCase()], [t(f.note)]] : [[''], [''], [''], ['']]);
}

/* Bascule une semaine dans l'application : reprend une fois les créneaux et les leçons de l'onglet du fichier maître. */
function basculer_(lundi, par) {
  if (gerees_()[lundi]) return false;
  var dates = datesSem_(lundi), sh = weekTab_(new Date(lundi + 'T12:00:00')), cren = [];
  if (sh) {
    var B = maitreBrut_(sh), mc = maitreCren_(B.v, dates);
    Object.keys(mc).forEach(function (k) { var c = mc[k]; if (!vide_(c.quoi)) cren.push({ date: k.split('|')[0], slot: k.split('|')[1], h: c.h, lieu: c.lieu, quoi: c.quoi, note: c.note }); });
    // leçons du fichier maître absentes de l'application → ajoutées comme validées
    var lp = leconsPlanifiees_(lundi), vus = {}, nv = [], valider = [];
    lp.forEach(function (x) { vus[cleL_(x.date, x.coach, x.h, x.tireur)] = x; });
    maitreLecons_(B, dates).forEach(function (m) {
      var ex = vus[cleL_(m.date, m.coach, m.h, m.tireur)];
      if (ex) { if (!/valid/i.test(ex.statut)) { ex.statut = 'validée'; valider.push(ex); } return; }
      nv.push({ date: m.date, h: m.h, duree: 30, coach: m.coach, tireur: m.tireur, vs: m.vs, statut: 'validée', note: m.entretien ? 'Entretien' : 'Repris du fichier maître' });
    });
    if (nv.length || valider.length) saveLecons_(lundi, nv.concat(valider), [], true);
  } else {
    crenType_().forEach(function (t) { if (t.j >= 0 && t.j < 6) cren.push({ date: dates[t.j], slot: t.slot, h: t.h, lieu: t.lieu, quoi: t.quoi, note: t.note }); });
  }
  crenRemplace_(lundi, cren, par);
  gereSh_().appendRow([lundi, new Date(), par || '']);
  _gerees = null;
  if (sh) PropertiesService.getScriptProperties().deleteProperty('MW_' + lundi);
  return true;
}

/* Recopie la semaine (créneaux + leçons validées) dans l'onglet du fichier maître, après avoir repris
 * ce qui y a été modifié à la main depuis la dernière écriture. Renvoie les éventuelles leçons qui n'ont pas trouvé de place. */
function ecrireMaitre_(lundi) {
  var sh = weekTab_(new Date(lundi + 'T12:00:00')); if (!sh || !gerees_()[lundi]) return { ok: false };
  var dates = datesSem_(lundi), P = PropertiesService.getScriptProperties(), K = 'MW_' + lundi, snap = null;
  try { snap = JSON.parse(P.getProperty(K) || 'null'); } catch (e) { }
  var B = maitreBrut_(sh), mc = maitreCren_(B.v, dates), ml = maitreLecons_(B, dates);
  var repris = 0;
  if (snap) {
    // 1. créneaux modifiés à la main dans le fichier maître
    Object.keys(mc).forEach(function (k) {
      var s = mc[k], sg = sigC_(s);
      if (sg !== (snap.c[k] || '')) { crenUpsert_(lundi, k.split('|')[0], k.split('|')[1], vide_(s.quoi) ? null : s, 'fichier maître'); repris++; }
    });
    // 2. leçons ajoutées ou effacées à la main dans le fichier maître
    var lp0 = leconsPlanifiees_(lundi), app = {}, mk = {}, sl = {}, nv = [], del = [];
    lp0.forEach(function (x) { app[cleL_(x.date, x.coach, x.h, x.tireur)] = x; });
    (snap.l || []).forEach(function (k) { sl[k] = 1; });
    ml.forEach(function (m) {
      var k = cleL_(m.date, m.coach, m.h, m.tireur); mk[k] = 1;
      if (sl[k]) return;
      var ex = app[k];
      if (ex) { if (!/valid/i.test(ex.statut)) { ex.statut = 'validée'; nv.push(ex); } return; }
      nv.push({ date: m.date, h: m.h, duree: 30, coach: m.coach, tireur: m.tireur, vs: m.vs, statut: 'validée', note: m.entretien ? 'Entretien' : 'Ajoutée dans le fichier maître' });
    });
    Object.keys(sl).forEach(function (k) { if (!mk[k] && app[k] && /valid/i.test(app[k].statut)) del.push(app[k].id); });
    if (nv.length || del.length) { saveLecons_(lundi, nv, del, true); repris += nv.length + del.length; }
  }
  // 3. écriture des créneaux
  var bc = crenBase_(lundi), ecrit = 0, nsnap = { c: {}, l: [] };
  for (var j = 0; j < 6; j++) SLOTS_.forEach(function (s) {
    var k = dates[j] + '|' + s, cible = bc[k] || null;
    if (cible && cible.note) { var nm = noteMaitre_(sh, j, s, cible.note); if (nm !== cible.note) cible = { h: cible.h, lieu: cible.lieu, quoi: cible.quoi, note: nm }; }
    nsnap.c[k] = sigC_(cible);
    var m = mc[k], egal = sigC_(cible) === sigC_(m) || (!cible && vide_(m.quoi) && !String(m.h || '').trim() && (s.charAt(0) === 'i' || !String(m.lieu || '').trim()));
    if (!egal) { ecrireSlot_(sh, j, s, cible); ecrit++; }
  });
  // 4. écriture de la grille des leçons validées
  var lp = leconsPlanifiees_(lundi).filter(function (x) { return /valid/i.test(x.statut); }), horsGrille = [];
  blocs_(B.L).forEach(function (bl) {
    var rows = 1 + 2 * bl.n, cur = B.v.slice(bl.coach, bl.coach + rows).map(function (r) { return r.slice(1, 25); });
    var tgt = cur.map(function (r, i) { return i === 0 ? r.slice() : r.map(function () { return ''; }); });
    var verts = [];
    for (var j2 = 0; j2 < 6; j2++) {
      var L = lp.filter(function (x) { return x.date === dates[j2] && (hmin_(x.h) < 13 * 60 ? 'matin' : 'après-midi') === bl.moment; })
        .sort(function (a, b) { return hmin_(a.h) - hmin_(b.h); });
      if (!L.length) continue;
      var heads = tgt[0].slice(4 * j2, 4 * j2 + 4), parC = {}, ordre = [];
      L.forEach(function (x) { var c = cc_(x.coach); if (!parC[c]) { parC[c] = []; ordre.push(c); } parC[c].push(x); });
      var colDe = {}, libres = [0, 1, 2, 3].filter(function (q) { return ordre.indexOf(cc_(heads[q])) < 0; });
      ordre.forEach(function (c) { var q = heads.map(cc_).indexOf(c); if (q < 0) { q = libres.shift(); if (q == null) return; tgt[0][4 * j2 + q] = parC[c][0].coach; } colDe[c] = q; });
      ordre.forEach(function (c) {
        var q = colDe[c], i = 0;
        parC[c].forEach(function (x) {
          var place = x.vs ? 2 : 1;
          if (q == null || i + place > bl.n) { horsGrille.push(x); return; }
          var col = 4 * j2 + q;
          tgt[1 + 2 * i][col] = x.h; tgt[2 + 2 * i][col] = String(x.tireur).toUpperCase();
          if (/entretien/i.test(x.note)) verts.push([2 + 2 * i, col]);
          i++;
          if (x.vs) { tgt[1 + 2 * i][col] = 'VS'; tgt[2 + 2 * i][col] = String(x.vs).toUpperCase(); i++; }
        });
      });
    }
    var diff = tgt.some(function (r, i) { return r.some(function (x, c) { return String(x) !== String(cur[i][c]); }); });
    var colsF = B.fc.slice(bl.coach, bl.coach + rows).map(function (r) { return r.slice(1, 25); }), vk = {};
    verts.forEach(function (p) { vk[p[0] + ',' + p[1]] = 1; });
    var diffF = false, nf = colsF.map(function (r, i) { return r.map(function (col, c) { var w = !!vk[i + ',' + c], is = vertC_(col); if (i > 0 && i % 2 === 0 && w !== is) { diffF = true; return w ? '#00b050' : '#000000'; } return col; }); });
    if (diff || diffF) {
      var rg = sh.getRange(bl.coach + 1, 2, rows, 24);
      if (diff) rg.setValues(tgt);
      if (diffF) rg.setFontColors(nf);
      ecrit++;
    }
  });
  lp.forEach(function (x) { if (horsGrille.indexOf(x) < 0) nsnap.l.push(cleL_(x.date, x.coach, x.h, x.tireur)); });
  P.setProperty(K, JSON.stringify(nsnap));
  if (ecrit) SpreadsheetApp.flush();
  return { ok: true, ecrit: ecrit, repris: repris, horsGrille: horsGrille.map(function (x) { return x.date + ' ' + x.h + ' ' + x.coach + ' · ' + x.tireur; }) };
}
/* Toutes les heures (appelé par notifSeances) : resynchronise les semaines gérées en cours et à venir. */
function syncMaitreAuto_() {
  var P = PropertiesService.getScriptProperties(), last = Number(P.getProperty('MW_last') || 0);
  if (Date.now() - last < 55 * 60000) return;
  P.setProperty('MW_last', String(Date.now()));
  var g = gerees_(), lock = LockService.getScriptLock(); if (!lock.tryLock(10000)) return;
  try {
    for (var o = -1; o <= 3; o++) { var l = fmt_(lundi_(o), 'yyyy-MM-dd'); if (g[l]) { var r = ecrireMaitre_(l); if (r.repris) { cacheClear_(); CacheService.getScriptCache().remove('saison_n'); } } }
  } finally { lock.releaseLock(); }
}
/* Semaine sans onglet dans le fichier maître : construite uniquement depuis l'application. */
function semaineApp_(lun) {
  var lundi = fmt_(lun, 'yyyy-MM-dd'), dates = datesSem_(lundi), g = gerees_()[lundi], bc = g ? crenBase_(lundi) : {}, ty = g ? [] : crenType_();
  var lp = leconsPlanifiees_(lundi).filter(function (x) { return /valid/i.test(x.statut); });
  var jours = dates.map(function (d, j) {
    var dd = new Date(d + 'T12:00:00'), cren = [];
    if (g) SLOTS_.forEach(function (s) { var c = bc[d + '|' + s]; if (c) cren.push({ h: c.h, lieu: c.lieu, quoi: c.quoi, note: c.note, slot: s }); });
    else ty.forEach(function (t) { if (t.j === j) cren.push({ h: t.h, lieu: t.lieu, quoi: t.quoi, note: t.note, slot: t.slot }); });
    cren.sort(function (a, b) { return hmin_(a.h) - hmin_(b.h); });
    var L = lp.filter(function (x) { return x.date === d; });
    return { jour: JOURS_[j] + ' ' + dd.getDate() + ' ' + MOIS_U[dd.getMonth()], date: d, evt: '', creneaux: cren, c: 1 + 4 * j, dimanche: j === 6,
      lecons: L.filter(function (x) { return !/entretien/i.test(x.note); }).map(lecApp_), entretiens: L.filter(function (x) { return /entretien/i.test(x.note); }).map(lecApp_) };
  });
  var k = isoWeek_(lun), cpt = compteCren_(jours);
  var sa = { onglet: null, gere: !!g, provisoire: !g, titre: 'PROGRAMME SEMAINE ' + k.w + ' - du ' + Number(dates[0].slice(8)) + ' ' + MOIS_U[lun.getMonth()].toLowerCase() + ' au ' + Number(dates[6].slice(8)) + ' ' + MOIS_U[new Date(dates[6] + 'T12:00:00').getMonth()].toLowerCase() + ' ' + dates[6].slice(0, 4),
    lundi: lundi, assauts: cpt.assauts, pp: cpt.pp, fonda: cpt.fonda, leconsPol: '', leconsPe: '', jours: jours, compets: [], suivantes: [] };
  return appliqueComp_(sa, dates);
}
function lecApp_(x) { return { moment: hmin_(x.h) < 13 * 60 ? 'matin' : 'après-midi', coach: x.coach, h: x.h, tireur: x.tireur, vs: x.vs, doublette: !!x.vs, id: x.id }; }
function compteCren_(jours) {
  var o = { assauts: 0, pp: 0, fonda: 0 };
  jours.forEach(function (J) { J.creneaux.forEach(function (c) { var q = String(c.quoi).toUpperCase(); if (/ASSAUT/.test(q)) o.assauts++; if (/PR[ÉE]PA|\bPP\b/.test(q)) o.pp++; if (/FONDA/.test(q)) o.fonda++; }); });
  return o;
}

/* ---- API : semaine type et copie de semaines ---- */
function getSemaineType(key) { staff_(key); return crenType_(); }
function enregistrerSemaineType(key, lundi) {
  staff_(key);
  var off = Math.round((new Date(lundi + 'T12:00:00') - lundi_(0)) / 6048e5), s = semaine_(off); if (!s) throw new Error('Semaine introuvable');
  var rows = [];
  s.jours.forEach(function (J, j) { if (j < 6) J.creneaux.forEach(function (c) { if (c.slot) rows.push([j, c.slot, c.h, c.lieu, c.quoi, c.note]); }); });
  var sh = typeSh_(); if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 6).setValues(rows);
  return crenType_();
}
/* source : 'type' (semaine type) ou le lundi d'une autre semaine. */
function copierCreneaux(key, lundi, source) {
  var u = staff_(key), dates = datesSem_(lundi), liste = [];
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    basculer_(lundi, u.nom);
    if (source === 'type') crenType_().forEach(function (t) { if (t.j >= 0 && t.j < 6) liste.push({ date: dates[t.j], slot: t.slot, h: t.h, lieu: t.lieu, quoi: t.quoi, note: t.note }); });
    else {
      var offS = Math.round((new Date(source + 'T12:00:00') - lundi_(0)) / 6048e5), s = semaine_(offS); if (!s) throw new Error('Semaine source introuvable');
      s.jours.forEach(function (J, j) { if (j < 6) J.creneaux.forEach(function (c) { if (c.slot) liste.push({ date: dates[j], slot: c.slot, h: c.h, lieu: c.lieu, quoi: c.quoi, note: c.note }); }); });
    }
    crenRemplace_(lundi, liste, u.nom);
    ecrireMaitre_(lundi);
  } finally { lock.releaseLock(); }
  cacheClear_(); CacheService.getScriptCache().remove('saison_n');
  return semaine_(Math.round((new Date(lundi + 'T12:00:00') - lundi_(0)) / 6048e5));
}
/* Lancement : bascule les semaines à partir de celle indiquée et initialise la semaine type. */
function lancerPlanningApp(depuis, nb, semaineTypeDepuis) {
  if (semaineTypeDepuis && !crenType_().length) {
    var dates = datesSem_(semaineTypeDepuis), sh = weekTab_(new Date(semaineTypeDepuis + 'T12:00:00')), mc = maitreCren_(maitreBrut_(sh).v, dates), rows = [];
    Object.keys(mc).forEach(function (k) { var c = mc[k]; if (!vide_(c.quoi)) rows.push([dates.indexOf(k.split('|')[0]), k.split('|')[1], c.h, c.lieu, c.quoi, c.note]); });
    rows.sort(function (a, b) { return a[0] - b[0] || hmin_(a[2]) - hmin_(b[2]); });
    if (rows.length) typeSh_().getRange(2, 1, rows.length, 6).setValues(rows);
  }
  var out = [];
  for (var i = 0; i < (nb || 1); i++) {
    var d = new Date(depuis + 'T12:00:00'); d.setDate(d.getDate() + 7 * i); var l = fmt_(d, 'yyyy-MM-dd');
    var b = basculer_(l, 'Guy'); var r = ecrireMaitre_(l); out.push(l + ' : ' + (b ? 'basculée' : 'déjà gérée') + ' · ' + JSON.stringify(r));
  }
  cacheClear_(); CacheService.getScriptCache().remove('saison_n');
  Logger.log(out.join('\n'));
  return out;
}
function getSemaine(key, offset) {
  var uc = user_(key);
  if (uc && (uc.role === 'Coach' || uc.role === 'Tireur')) { var sc = semaine_(offset || 0), lc = sc ? sc.lundi : fmt_(lundi_(offset || 0), 'yyyy-MM-dd'); return { semaine: sc, seances: [], lundi: lc, pp: pp_(lc), memos: [] }; }
  staff_(key);
  var s = semaine_(offset || 0);
  var lun = s ? s.lundi : fmt_(lundi_(offset || 0), 'yyyy-MM-dd');
  return { semaine: s, seances: seancesSemaine_(lun), lundi: lun, pp: pp_(lun), memos: memosSemaine_(lun) };
}

function saveSeance(key, f) {
  var u = staff_(key);
  var sh = sheet_(IDS.BASE, 'Séances');
  var id = f.id || (String(f.date).replace(/-/g, '') + '-' + String(f.h || '').replace(/[^\dh]/g, '') + '-' + norm_(f.type).split(' ')[0].slice(0, 5) + '-' + Math.floor(Math.random() * 900 + 100));
  var row = [id, new Date(f.date + 'T12:00:00'), f.h || '', Number(f.duree) || '', f.type || '', f.lieu || '', f.coach || '', f.consignes || '', f.intensite || '', f.repos || '', f.photo || '', f.notes || '', u.nom, new Date()];
  var ligne = trouveLigne_(sh, 1, id);
  if (ligne) sh.getRange(ligne, 1, 1, 12).setValues([row.slice(0, 12)]);
  else sh.appendRow(row);
  cacheClear_();
  return id;
}
function trouveLigne_(sh, col, val) {
  if (sh.getLastRow() < 2) return 0;
  var v = sh.getRange(2, col, sh.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === String(val)) return i + 2;
  return 0;
}

/* Appel : on remplace toutes les lignes de la séance par la liste reçue. */
function saveAppel(key, id, date, liste) {
  var u = staff_(key);
  var sh = sheet_(IDS.BASE, 'Appel');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (sh.getLastRow() > 1) {
      var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
      for (var i = v.length - 1; i >= 0; i--) if (v[i][0] === String(id)) sh.deleteRow(i + 2);
    }
    var d = new Date(date + 'T12:00:00'), now = new Date();
    var out = liste.filter(function (x) { return x.statut; }).map(function (x) { return [id, d, x.tireur, x.statut, x.retard || '', x.duree || '', x.motif || '', now, u.nom]; });
    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, 9).setValues(out);
  } finally { lock.releaseLock(); }
  return 'ok';
}

function saveExercice(key, f) {
  staff_(key);
  var sh = sheet_(IDS.BASE, 'Exercices');
  var idE = f.idE || (f.idS + '-E' + Math.floor(Math.random() * 9000 + 1000));
  if (sh.getLastColumn() < 12) sh.getRange(1, 11, 1, 2).setValues([['Matchs en 5 touches', 'Matchs en 15 touches']]);
  var row = [f.idS, idE, f.ordre || '', f.format || '', f.touches || '', f.nb || '', f.intensite || '', f.repos || '', f.consignes || '', (f.participants || []).join(', '), f.n5 || '', f.n15 || ''];
  var ligne = trouveLigne_(sh, 2, idE);
  if (ligne) sh.getRange(ligne, 1, 1, 12).setValues([row]); else sh.appendRow(row);
  return idE;
}

/* Matchs d'entraînement (poule, tableau, équipes, libres) → onglet Assauts. */
function saveMatchs(key, idS, idE, date, format, type, matchs) {
  var u = staff_(key);
  var sh = sheet_(IDS.BASE, 'Assauts');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (idE && sh.getLastRow() > 1) {
      var v = sh.getRange(2, 12, sh.getLastRow() - 1, 1).getDisplayValues();
      for (var i = v.length - 1; i >= 0; i--) if (v[i][0] === String(idE)) sh.deleteRow(i + 2);
    }
    var d = date ? new Date(date + 'T12:00:00') : new Date();
    var out = matchs.filter(function (m) { return m.a && m.b && m.sa !== '' && m.sb !== '' && m.sa != null && m.sb != null; })
      .map(function (m) { return [d, type || '', m.a, m.b, Number(m.sa), Number(m.sb), '', format || '', u.nom, type || '', idS || '', idE || '']; });
    if (out.length) {
      var start = sh.getLastRow() + 1;
      sh.getRange(start, 1, out.length, 6).setValues(out.map(function (r) { return r.slice(0, 6); }));
      sh.getRange(start, 8, out.length, 5).setValues(out.map(function (r) { return r.slice(7, 12); }));
    }
  } finally { lock.releaseLock(); }
  cacheClear_();
  return out.length;
}

function dossierApp_(sous) {
  var it = DriveApp.getFoldersByName(DOSSIER_APP), dir = it.hasNext() ? it.next() : DriveApp.createFolder(DOSSIER_APP);
  if (!sous) return dir;
  var it2 = dir.getFoldersByName(sous);
  return it2.hasNext() ? it2.next() : dir.createFolder(sous);
}
function uploadPhoto(key, idS, nomFichier, dataUrl) {
  staff_(key);
  var m = String(dataUrl).match(/^data:(.+?);base64,(.*)$/);
  if (!m) throw new Error('Image illisible');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (idS || 'seance') + ' · ' + (nomFichier || 'tableau.jpg'));
  var f = dossierApp_('Photos séances').createFile(blob);
  var url = f.getUrl();
  if (idS) {
    var sh = sheet_(IDS.BASE, 'Séances'), l = trouveLigne_(sh, 1, idS);
    if (l) { var old = sh.getRange(l, 11).getValue(); sh.getRange(l, 11).setValue(old ? old + '\n' + url : url); }
  }
  return url;
}

/* =========================================================
 * Résultats : cache Drive construit depuis V6
 * ======================================================= */

function construireCacheResultats() {
  var dir = dossierApp_('Cache');
  CATS.forEach(function (cat) {
    var sh = sheet_(IDS.RESULTATS, cat);
    if (!sh) return;
    var nr = sh.getLastRow(), nc = sh.getLastColumn();
    var v = sh.getRange(1, 1, nr, nc).getDisplayValues();
    var groupes = [], cur = null;
    for (var c = 0; c < nc; c++) {
      var code = v[0][c] || (cur ? cur.code : '');
      if (!cur || code !== cur.code) { cur = { code: code, titre: v[1][c] || '', cols: [] }; groupes.push(cur); }
      if (!cur.titre && v[1][c]) cur.titre = v[1][c];
      cur.cols.push(c);
    }
    var lignes = [];
    for (var r = 4; r < nr; r++) {
      if (!v[r][2]) continue;
      var row = v[r].slice(), last = row.length;
      while (last > 0 && row[last - 1] === '') last--;
      lignes.push(row.slice(0, last));
    }
    var data = { cat: cat, maj: fmt_(new Date(), 'dd/MM/yyyy HH:mm'), titres: v[2], sous: v[3], groupes: groupes, lignes: lignes };
    var nom = 'resultats_' + cat.replace(/\s/g, '_') + '.json', it = dir.getFilesByName(nom);
    var contenu = JSON.stringify(data);
    if (it.hasNext()) it.next().setContent(contenu); else dir.createFile(nom, contenu, 'application/json');
  });
  CacheService.getScriptCache().removeAll(CATS.map(function (c) { return 'res_' + c.replace(/\s/g, '_') + '_n'; }));
  return 'ok';
}

function getResultats(key, cat) {
  var u = user_(key);
  if (!u) throw new Error('Lien invalide');
  if (CATS.indexOf(cat) < 0) cat = CATS[1];
  var ck = 'res_' + cat.replace(/\s/g, '_'), d = cacheGet_(ck);
  if (!d) {
    var it = dossierApp_('Cache').getFilesByName('resultats_' + cat.replace(/\s/g, '_') + '.json');
    if (!it.hasNext()) { construireCacheResultats(); it = dossierApp_('Cache').getFilesByName('resultats_' + cat.replace(/\s/g, '_') + '.json'); }
    d = JSON.parse(it.next().getBlob().getDataAsString());
    cachePut_(ck, d, 3600);
  }
  d.pole = tireurs_().filter(function (t) { return t.pole; }).map(function (t) { return t.cle; });
  try { d.part = profilsLecons_().filter(function (p) { return /PARTENAIRE/i.test(p.groupe); }).map(function (p) { return [p.nom, p.prenom || '']; }); } catch (e) { d.part = []; }
  return d;
}

/* Tableau imprimable « classement sélectif + nationaux » : onglets IMPR du fichier résultats (V6). */
function getImpr(key, cat) {
  if (!user_(key)) throw new Error('Lien invalide');
  var ck = 'impr_' + norm_(cat).replace(/ /g, '_'), c = cacheGet_(ck);
  if (c) return c;
  var sh = sheet_(IDS.RESULTATS, 'IMPR ' + cat);
  if (!sh) return null;
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var heads = v[2], n = heads.length; while (n > 0 && !heads[n - 1]) n--;
  var groupes = [], g = null;
  for (var j = 0; j < n; j++) { if (v[1][j]) { g = { label: v[1][j], start: j, span: 1 }; groupes.push(g); } else if (g) g.span++; }
  var rows = v.slice(3).filter(function (r) { return r[2]; }).map(function (r) { return r.slice(0, n); });
  var notes = v.slice(3).filter(function (r) { return !r[2] && r[0] && String(r[0]).length > 30; }).map(function (r) { return r[0]; });
  var out = { titre: v[0][0], groupes: groupes, heads: heads.slice(0, n), rows: rows, notes: notes };
  cachePut_(ck, out, 1800);
  return out;
}
/* ---- Import des classements nationaux Senior (réel, sans ER recalculé, M23) depuis un fichier JSON du Drive ----
 * Format : { FH: "code,rang,rangSansER,rangM23;...", FD: "..." }. Met à jour les colonnes 2026/2027 des onglets de V6. */
function evolTxt_(o, n) {
  if (n === '' || n == null) return '';
  if (o === '' || o == null || isNaN(Number(o))) return 'Nouveau';
  var d = Number(o) - Number(n);
  if (d === 0) return '= 0'; if (d >= 20) return '🚀 +' + d; if (d >= 10) return '↗ +' + d; if (d > 0) return '▲ +' + d; if (d <= -10) return '⚠ ' + d; return '▼ ' + d;
}
function majClassementsSenior(key, fileId, libelle) {
  staff_(key);
  var data = JSON.parse(DriveApp.getFileById(fileId).getBlob().getDataAsString()), rep = [];
  ['FH', 'FD'].forEach(function (sx) {
    var R = {}; String(data[sx] || '').split(';').forEach(function (l) { var p = l.split(','); if (p[0]) R[p[0]] = { r: p[1], s: p[2], m: p[3], o: p.length > 4 ? p[4] : null }; });
    ['M17 ', 'M20 ', 'M23 ', 'Senior '].forEach(function (t) {
      var sh = sheet_(IDS.RESULTATS, t + sx); if (!sh) return;
      var lc = sh.getLastColumn(), lr = sh.getLastRow(); if (lr < 5) return;
      var h = sh.getRange(1, 1, 3, lc).getDisplayValues(), codes = sh.getRange(5, 2, lr - 4, 1).getDisplayValues().map(function (x) { var c = String(x[0]).trim(); return /^\d+$/.test(c) ? ('000000' + c).slice(-6) : c; });
      function col(hd) { for (var c = 0; c < lc; c++) if (String(h[2][c]).replace(/\n/g, ' ').trim() === hd) return c + 1; return 0; }
      var cR = col('Clt Senior 26/27'), cS = col('Clt Senior sans ER 26/27'), cE = col('Évol Senior sans ER vs import précédent'), cM = col('Clt M23 26/27'), cME = col('Évol M23 26/27 vs 25/26'), cM25 = col('Clt M23 25/26');
      function put(c, vals) { if (c) sh.getRange(5, c, vals.length, 1).setValues(vals.map(function (v) { return [v]; })); }
      var n = 0;
      if (cS && cE) { var old = sh.getRange(5, cS, codes.length, 1).getValues(); put(cE, codes.map(function (k, i) { return R[k] ? evolTxt_(R[k].o != null ? R[k].o : old[i][0], R[k].s) : ''; })); }
      if (cR) { put(cR, codes.map(function (k) { if (R[k]) n++; return R[k] ? Number(R[k].r) : ''; })); }
      if (cS) put(cS, codes.map(function (k) { return R[k] && R[k].s ? Number(R[k].s) : ''; }));
      if (cM) { put(cM, codes.map(function (k) { return R[k] && R[k].m ? Number(R[k].m) : ''; }));
        if (cME && cM25) { var o25 = sh.getRange(5, cM25, codes.length, 1).getValues(); put(cME, codes.map(function (k, i) { return R[k] && R[k].m ? evolTxt_(o25[i][0], R[k].m) : ''; })); } }
      rep.push([t + sx, 'Senior' + (cM ? ' + M23' : ''), n]);
    });
  });
  var rp = sheet_(IDS.RESULTATS, 'Rapport import');
  if (rp) rp.getRange(rp.getLastRow() + 2, 1, rep.length + 1, 3).setValues([['Import ' + (libelle || 'classements Senior') + ' · ' + fmt_(new Date(), 'dd/MM/yyyy HH:mm'), '', '']].concat(rep));
  construireCacheResultats();
  return rep;
}
/* ---- Écriture de places dans V6 depuis un fichier JSON du Drive ----
 * { items: [{ sx, code: 'RESULTATS_INTER|Senior|2026/2027', match: 'EI 1', titre, note, places: { licence: place } }], src: [lignes SRC_RESULTS] } */
function majPlaces(key, fileId) {
  staff_(key);
  var data = JSON.parse(DriveApp.getFileById(fileId).getBlob().getDataAsString()), rep = [];
  (data.items || []).forEach(function (it) {
    ['M17 ', 'M20 ', 'M23 ', 'Senior '].forEach(function (t) {
      var sh = sheet_(IDS.RESULTATS, t + it.sx); if (!sh) return;
      var lc = sh.getLastColumn(), lr = sh.getLastRow(); if (lr < 5) return;
      var h = sh.getRange(1, 1, 3, lc).getDisplayValues(), col = 0;
      for (var c = 0; c < lc && !col; c++) { var hd = String(h[2][c]).replace(/\n/g, ' ').trim(); if (String(h[0][c]) === it.code && (hd === it.match || hd.indexOf(it.match + ' ') === 0) && !/^Pts/i.test(hd)) col = c + 1; }
      if (!col) return;
      var hd0 = String(h[2][col - 1]).trim(); if (hd0 === it.match && it.titre) sh.getRange(3, col).setValue(it.titre);
      var codes = sh.getRange(5, 2, lr - 4, 1).getDisplayValues().map(function (x) { var c = String(x[0]).trim(); return /^\d+$/.test(c) ? ('000000' + c).slice(-6) : c; });
      var rg = sh.getRange(5, col, codes.length, 1), v = rg.getValues(), nt = rg.getNotes(), n = 0;
      codes.forEach(function (k, i) { if (it.places[k] != null) { v[i][0] = Number(it.places[k]); if (it.note && !nt[i][0]) nt[i][0] = it.note; n++; } });
      if (n) { rg.setValues(v); rg.setNotes(nt); }
      rep.push([t + it.sx, it.titre || it.match, n]);
    });
  });
  if (data.src && data.src.length) {
    var sr = sheet_(IDS.RESULTATS, 'SRC_RESULTS'), lr2 = sr.getLastRow(), ex = {};
    if (lr2 > 1) sr.getRange(2, 1, lr2 - 1, 13).getDisplayValues().forEach(function (x) { ex[x[8] + '|' + x[2] + '|' + x[4]] = 1; });
    var add = data.src.filter(function (x) { return !ex[x[8] + '|' + x[2] + '|' + x[4]]; });
    if (add.length) sr.getRange(lr2 + 1, 1, add.length, 13).setNumberFormat('@').setValues(add);
    rep.push(['SRC_RESULTS', 'lignes ajoutées', add.length]);
  }
  var rp = sheet_(IDS.RESULTATS, 'Rapport import');
  if (rp) rp.getRange(rp.getLastRow() + 2, 1, rep.length + 1, 3).setValues([['Places · ' + fmt_(new Date(), 'dd/MM/yyyy HH:mm'), '', '']].concat(rep));
  construireCacheResultats();
  return rep;
}
function actualiserResultats(key) { staff_(key); construireCacheResultats(); return 'ok'; }

/* Parcours détaillés (poules + tableau) : onglet SRC_PARCOURS de V6. */
function dateNum_(s) { var m = String(s).match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? Number(m[3] + m[2] + m[1]) : 0; }
function parcours_(nom) {
  var k = norm_(nom), sh = sheet_(IDS.RESULTATS, 'SRC_PARCOURS');
  if (!sh || sh.getLastRow() < 2) return [];
  var out = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues()
    .filter(function (x) { return x[3] && norm_(x[3]) === k; })
    .map(function (x) { return { ep: x[0], date: x[1], place: x[4], txt: x[5], t: dateNum_(x[1]) }; });
  out.sort(function (a, b) { return b.t - a.t; });
  return out;
}

/* =========================================================
 * Fiche tireur
 * ======================================================= */

function fiche_(nom) {
  var t = trouveTireur_(nom), k = t ? t.cle : norm_(nom);
  function moi(n) { var x = norm_(n); return x === k || (t && (x === norm_(t.nom) || x === norm_(t.prenom + ' ' + t.nom))); }
  var a = assauts_();
  var mat = materiel_();
  var lames = mat.lames.filter(function (x) { return moi(x.tireur); });
  var dettes = dettes_().filter(function (x) { return moi(x.tireur); });
  lames.forEach(function (l) { if (l.statut === 'Vendue' && !l.paye) dettes.push({ date: l.date, lib: 'Lame ' + l.n + (l.couleur ? ' (' + l.couleur + ')' : ''), montant: num_(l.prix), paye: false, lame: true }); });
  var tests = tests_(), mesTests = tests.filter(function (x) { return moi(x.tireur); });
  var moyennes = {};
  tests.forEach(function (x) { var c = x.test + '|' + x.cote; (moyennes[c] = moyennes[c] || []).push(x.val); });
  Object.keys(moyennes).forEach(function (c) { var l = moyennes[c]; moyennes[c] = l.reduce(function (s, v) { return s + v; }, 0) / l.length; });
  return {
    tireur: t, nom: nom, libs: LIB_FICHE,
    resultats: resultatsResume_()[k] || (t ? resultatsResume_()[norm_(String(t.nom).split(/[\s-]+/)[0] + ' ' + t.prenom)] || resultatsResume_()['NOM ' + norm_(t.nom)] || resultatsResume_()['NI ' + norm_(t.nom) + ' ' + norm_(t.prenom).charAt(0)] : null) || {},
    parcours: parcours_(t ? t.nom + ' ' + t.prenom : nom),
    wellness: wellness_(30).filter(function (x) { return moi(x.tireur); }),
    assauts: a.parTireur[k] || a.parTireur[t ? norm_(t.nom) : ''] || null,
    matchs: assautsBruts_().filter(function (m) { return moi(m.a) || moi(m.b); }).slice(-30).reverse(),
    lames: lames, dettes: dettes, tests: mesTests, moyennes: moyennes,
    scolaire: scolaire_().filter(function (x) { return moi(x.tireur); }),
    presences: presences_(moi)
  };
}
function presences_(moi) {
  var sh = sheet_(IDS.BASE, 'Appel');
  if (sh.getLastRow() < 2) return {};
  var out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (x) { if (moi(x[2])) out[x[3]] = (out[x[3]] || 0) + 1; });
  return out;
}

function getFiche(key, nom) {
  var u = user_(key);
  if (!u || u.role === 'Invité') throw new Error('Accès non autorisé');
  if (u.role === 'Tireur' && norm_(nom) !== norm_(u.nom)) throw new Error('Accès réservé à ta propre fiche');
  var f = fiche_(nom);
  if (u.role === 'Staff') { var k = norm_(nom); f.memos = memos_().filter(function (m) { return m.tireurs.some(function (x) { return norm_(x) === k; }); }).reverse(); f.suivi = suivi_().filter(function (x) { return norm_(x.tireur) === k; }).reverse(); }
  return f;
}

/* Le tireur propose une modification ; le staff valide. Le staff écrit directement. */
function saveChampFiche(key, nom, champ, valeur) {
  var u = user_(key);
  if (!u || u.role === 'Invité') throw new Error('Accès non autorisé');
  var t = trouveTireur_(u.role === 'Tireur' ? u.nom : nom);
  if (!t) throw new Error('Tireur introuvable dans l\'onglet Tireurs');
  var col = CHAMPS_FICHE.indexOf(champ) + 1;
  if (col < 1) throw new Error('Champ inconnu');
  if (u.role === 'Staff') { sheet_(IDS.BASE, 'Tireurs').getRange(t.ligne, col).setValue(valeur); cacheClear_(); return 'enregistré'; }
  sheet_(IDS.BASE, 'Demandes fiche').appendRow([new Date(), t.nom + ' ' + t.prenom, champ, t[champ] || '', valeur, 'À valider', '']);
  return 'envoyé au staff';
}
function traiterDemande(key, ligne, accepter) {
  var u = staff_(key), sh = sheet_(IDS.BASE, 'Demandes fiche');
  var x = sh.getRange(ligne, 1, 1, 7).getValues()[0];
  if (accepter) {
    var t = trouveTireur_(x[1]), col = CHAMPS_FICHE.indexOf(x[2]) + 1;
    if (t && col > 0) sheet_(IDS.BASE, 'Tireurs').getRange(t.ligne, col).setValue(x[4]);
  }
  sh.getRange(ligne, 6, 1, 2).setValues([[accepter ? 'Validé' : 'Refusé', u.nom]]);
  cacheClear_();
  return 'ok';
}

function saveTest(key, f) {
  var u = staff_(key);
  sheet_(IDS.BASE, 'Tests physiques').appendRow([new Date(f.date + 'T12:00:00'), f.tireur, f.famille || '', f.test, Number(String(f.val).replace(',', '.')), f.unite || '', f.cote || '-', f.haut === 'non' ? 'non' : 'oui', u.nom, f.rq || '']);
  return 'ok';
}
function saveDette(key, f) {
  staff_(key);
  sheet_(IDS.BASE, 'Dettes tireurs').appendRow([new Date(), f.tireur, f.lib, Number(String(f.montant).replace(',', '.')), false, '', '', f.note || '']);
  return 'ok';
}
function regleDette(key, ligne, moyen) {
  staff_(key);
  sheet_(IDS.BASE, 'Dettes tireurs').getRange(ligne, 5, 1, 3).setValues([[true, new Date(), moyen || '']]);
  return 'ok';
}

/* =========================================================
 * API principale
 * ======================================================= */

function bundleStaff_() {
  var b = cacheGet_('bundle');
  if (b) return b;
  var s = semaine_(0);
  b = {
    semaine: s, pp: s ? pp_(s.lundi) : [], vacations: vacations_(), budget: budget_(), materiel: materiel_(), textile: textile_(),
    tireurs: tireurs_(), resultats: resultatsResume_(), wellness: wellness_(42), assauts: assauts_(),
    scolaire: scolaire_(), demandes: demandes_(), cats: CATS,
    dettes: dettes_().filter(function (x) { return !x.paye && x.montant; }).map(function (x) { return { tireur: x.tireur, lib: x.lib, montant: x.montant }; }),
    rpe: rpe_(63),
    maj: fmt_(new Date(), 'dd/MM/yyyy HH:mm')
  };
  cachePut_('bundle', b, 300);
  return b;
}

function getData(key) {
  var u = user_(key);
  if (!u) return { erreur: 'Lien invalide. Demande ton lien personnel à Guy.' };
  vu_(key);
  if (u.role === 'Tireur') { var sT = null; try { sT = semaine_(0); } catch (e) { } return { user: u, fiche: fiche_(u.nom), cats: CATS, semaine: sT, pp: sT ? pp_(sT.lundi) : [] }; }
  if (u.role === 'Sélectionneur') return { user: u, cats: CATS };
  if (u.role === 'Coach') { var sC = semaine_(0); return { user: u, cats: CATS, semaine: sC, pp: sC ? pp_(sC.lundi) : [], tireurs: tireurs_().map(function (t) { return { nom: t.nom, prenom: t.prenom, pole: t.pole, statut: t.statut, cle: t.cle }; }) }; }
  if (u.role !== 'Staff') return { user: u, cats: CATS };
  u.admin = ADMIN.indexOf(u.nom) >= 0;
  var b = bundleStaff_(); b.user = u;
  return b;
}

function rafraichir(key) { staff_(key); cacheClear_(); return getData(key); }

/* =========================================================
 * Saisies existantes
 * ======================================================= */

function saveWellness(key, f) {
  var u = user_(key);
  if (!u || u.role === 'Invité') throw new Error('Lien invalide');
  var nom = u.role === 'Tireur' ? u.nom : f.tireur;
  var vals = [f.s, f.f, f.c, f.hu, f.st].map(Number);
  if (vals.some(function (v) { return !(v >= 1 && v <= 5); })) throw new Error('Chaque note doit être entre 1 et 5');
  var sh = sheet_(IDS.BASE, 'Wellness'), row = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(row, 1, 1, 7).setValues([[new Date(), nom].concat(vals)]);
  sh.getRange(row, 9).setValue(f.comm || '');
  cacheClear_();
  /* notification immédiate à Guy (plus dans le résumé de 12 h / 21 h) */
  if (u.role === 'Tireur') try {
    var tot = vals.reduce(function (a, b) { return a + b; }, 0);
    pushEnvoyer_(SUIVI_ADMIN, (tot <= 13 ? '⚠ ' : '') + 'Wellness · ' + prenomDe_(nom) + ' ' + tot + '/25',
      'Sommeil ' + vals[0] + ' · forme ' + vals[1] + ' · courbatures ' + vals[2] + ' · humeur ' + vals[3] + ' · stress ' + vals[4] + (f.comm ? ' — ' + f.comm : ''), '#jour');
  } catch (e) { }
  return 'ok';
}
function saveAssaut(key, f) {
  if (!f.a || !f.b || f.a === f.b) throw new Error('Choisis deux tireurs différents');
  return saveMatchs(key, f.idS || '', '', f.date || '', f.format || '', f.seance || 'Assauts', [{ a: f.a, b: f.b, sa: f.sa, sb: f.sb }]);
}
function saveLame(key, f) {
  staff_(key);
  var sh = sheet_(IDS.BASE, 'Lames'), l = trouveLigne_(sh, 1, f.n);
  if (!l) throw new Error('Lame introuvable');
  if (f.couleur) sh.getRange(l, 2).setValue(f.couleur);
  sh.getRange(l, 5, 1, 4).setValues([[f.statut, f.tireur || '', new Date(), !!f.paye]]);
  cacheClear_();
  return 'ok';
}
function saveScolaire(key, f) {
  var u = user_(key);
  if (!u || u.role === 'Invité') throw new Error('Lien invalide');
  sheet_(IDS.BASE, 'Scolaire').appendRow([u.role === 'Tireur' ? u.nom : f.tireur, f.date ? new Date(f.date) : '', f.matiere || '', f.type || '', f.besoin || '', 'À faire', u.nom]);
  cacheClear_();
  return 'ok';
}

/* =========================================================
 * v5 : textile, prépa physique, participants, présences, parcours au survol
 * ======================================================= */

function isoTxt_(s) { var m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : String(s).slice(0, 10); }
function textile_() {
  return rows_(IDS.BASE, 'Dotation textile').slice(1).filter(function (x) { return x[1] || x[2]; })
    .map(function (x) { return { date: x[0], tireur: x[1], article: x[2], taille: x[3], qte: x[4], statut: x[5], origine: x[6], note: x[7] }; });
}
function saveTextile(key, f) {
  staff_(key);
  if (!f.article) throw new Error('Indique l\'article');
  sheet_(IDS.BASE, 'Dotation textile').appendRow([new Date(), f.tireur || '', f.article, f.taille || '', Number(f.qte) || 1, f.statut || 'À distribuer', f.origine || '', f.note || '']);
  cacheClear_();
  return 'ok';
}
function pp_(lundiIso) {
  return rows_(IDS.BASE, 'Prépa physique (Mathieu)').slice(1).filter(function (x) { return isoTxt_(x[0]) === lundiIso; })
    .map(function (x) { return { bloc: x[1], ordre: x[2], txt: x[3] }; });
}
/* =========================================================
 * Compétitions et priorités gérées dans l'application (étape 2 du détachement du fichier maître)
 * BASE › Compétitions (ID + mêmes colonnes que « Liste compet »), Priorités (ID × tireur), Calendrier tireurs (colonnes du calendrier).
 * Chaque modification est recopiée dans « Liste compet » et « Priorisation A-D » du fichier maître pendant la transition.
 * ======================================================= */
var _comps = null;
function compSh_() { var sh = ss_(IDS.BASE).getSheetByName('Compétitions'); if (!sh) throw new Error('Onglet Compétitions absent de BASE'); return sh; }
function comps_() {
  if (_comps) return _comps;
  var sh = compSh_(), n = sh.getLastRow() - 1, o = [];
  if (n > 0) sh.getRange(2, 1, n, 35).getDisplayValues().forEach(function (x, i) { if (x[0]) o.push({ id: x[0], v: x.slice(1), ligne: i + 2 }); });
  return (_comps = o);
}
function prios_() {
  var o = {};
  rows_(IDS.BASE, 'Priorités').slice(1).forEach(function (x) { if (x[0] && x[1]) (o[x[0]] = o[x[0]] || {})[x[1]] = String(x[2] || '').trim(); });
  return o;
}
function calTireurs_() {
  return rows_(IDS.BASE, 'Calendrier tireurs').slice(1).filter(function (x) { return x[1]; })
    .sort(function (a, b) { return Number(a[0]) - Number(b[0]); }).map(function (x) { return { nom: String(x[1]).trim(), tireur: x[2], groupe: x[3] }; });
}
function isoFr_(s) { var m = String(s || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : ''; }
function frIso_(s) { var p = String(s || '').split('-'); return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : ''; }
function compEv_(c, pr, heads) {
  var x = c.v, d1 = isoFr_(x[0]);
  return { id: c.id, hf: x[3], d1: d1, d2: isoFr_(x[1]) || d1, dep: isoFr_(x[10]), ret: isoFr_(x[11]), nom: String(x[6]).trim(), lieu: x[7], cat: x[4], ie: x[5], type: x[8], cal: x[9],
    chd: x[12], ent: x[13], arb: x[14], deleg: x[15], hotel: x[16], tr: [x.slice(17, 21), x.slice(21, 25), x.slice(25, 29), x.slice(29, 33)].map(function (t) { return { d: isoFr_(t[0]), v: t[1], p: t[2], n: t[3] }; }), note: x[33],
    p: heads.map(function (h) { return (pr[c.id] || {})[h.nom] || ''; }) };
}
function getCalendrier(key) {
  staff_(key);
  var c = cacheGet_('cal');
  if (c) return c;
  var heads = calTireurs_(), pr = prios_(), ev = [];
  comps_().forEach(function (x) { if (String(x.v[3]).indexOf('H') >= 0 && isoFr_(x.v[0])) ev.push(compEv_(x, pr, heads)); });
  ev.sort(function (a, b) { return a.d1 < b.d1 ? -1 : a.d1 > b.d1 ? 1 : 0; });
  var out = { heads: heads, ev: ev, evts: evenements_() };
  cachePut_('cal', out, 1800);
  return out;
}
/* Tireur : les compétitions où il est prévu (priorités A, B, C ; D = à confirmer), sans afficher les lettres. */
function getMesCompets(key) {
  var u = user_(key); if (!u) throw new Error('Lien invalide');
  var heads = calTireurs_(), n = norm_(u.nom), t = trouveTireur_(u.nom), pre = t ? norm_(t.prenom) : '';
  var c = heads.map(function (h, k) { return { h: h, k: k }; }).filter(function (x) { var nm = norm_(x.h.tireur || '').split(' ')[0]; return nm && (' ' + n + ' ').indexOf(' ' + nm + ' ') >= 0; });
  if (c.length > 1) { var f = c.filter(function (x) { var p = norm_(String(x.h.nom).split(' ')[0]); return p && (pre || n).indexOf(p.slice(0, 3)) >= 0; }); if (f.length) c = f; }
  if (!c.length) return { trouve: false, ev: [] };
  var i = c[0].k, pr = prios_(), out = [];
  comps_().forEach(function (x) {
    if (String(x.v[3]).indexOf('H') < 0 || !isoFr_(x.v[0])) return;
    var e = compEv_(x, pr, heads), p = String(e.p[i] || '').toUpperCase();
    if (!/^[ABCD]$/.test(p)) return;
    out.push({ d1: e.d1, d2: e.d2, dep: e.dep, ret: e.ret, nom: e.nom, lieu: e.lieu, cat: e.cat, type: e.type, cal: e.cal, ie: e.ie, hotel: e.hotel, ent: e.ent, statut: p === 'D' ? 'À confirmer' : 'Prévue' });
  });
  out.sort(function (a, b) { return a.d1 < b.d1 ? -1 : 1; });
  return { trouve: true, lib: heads[i].nom, ev: out };
}
function getParticipants(key, dateTxt, lieu) {
  staff_(key);
  var m = String(dateTxt).match(/(\d{1,2})\/(\d{1,2})/), dd = m ? ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) : '', L = norm_(lieu);
  var heads = calTireurs_(), pr = prios_(), cs = comps_();
  for (var i = 0; i < cs.length; i++) {
    var x = cs[i].v, lx = norm_(x[7]);
    if (String(x[0]).slice(0, 5) !== dd || String(x[3]).indexOf('H') < 0) continue;
    if (L && lx && lx.indexOf(L) < 0 && L.indexOf(lx) < 0) continue;
    var p = pr[cs[i].id] || {};
    return { id: cs[i].id, nom: x[6], lieu: x[7], type: x[8], cal: x[9], cat: x[4], du: x[0], au: x[1], liste: heads.map(function (h) { return { nom: h.nom, prio: p[h.nom] || '', groupe: h.groupe }; }) };
  }
  return null;
}
function compParId_(id) { var c = comps_().filter(function (x) { return x.id === id; })[0]; if (!c) throw new Error('Compétition introuvable'); return c; }
function refMaitre_(c) { return { d1: isoFr_(c.v[0]), nom: String(c.v[6]).trim(), lieu: c.v[7] }; }
function savePriorite(key, ref, tireur, val) {
  staff_(key);
  val = String(val || '').trim();
  if (['A', 'B', 'C', 'D', 'non', ''].indexOf(val) < 0) throw new Error('Priorité inconnue');
  var c = ref.id ? compParId_(ref.id) : comps_().filter(function (x) { var r = refMaitre_(x); return r.d1 === ref.d1 && r.nom === String(ref.nom).trim() && String(r.lieu) === String(ref.lieu); })[0];
  if (!c) throw new Error('Compétition introuvable');
  var sh = sheet_(IDS.BASE, 'Priorités'), v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues() : [], l = 0;
  for (var i = 0; i < v.length; i++) if (v[i][0] === c.id && v[i][1] === tireur) { l = i + 2; break; }
  if (l) sh.getRange(l, 3, 1, 2).setValues([[val, new Date()]]); else sh.appendRow([c.id, tireur, val, new Date()]);
  CacheService.getScriptCache().remove('cal_n');
  try { // copie dans le fichier maître
    var shm = sheet_(IDS.MAITRE, 'Priorisation A-D'), lm = ligneCompet_(shm, refMaitre_(c));
    var head = shm.getRange(3, 1, 1, shm.getLastColumn()).getDisplayValues()[0], cm = head.indexOf(tireur);
    if (lm && cm >= 12) shm.getRange(lm, cm + 1).setValue(val);
  } catch (e) { Logger.log('Priorité → maître : ' + e); }
  return 'ok';
}
/* f : champs du formulaire ; ref : { id } de la compétition modifiée (absent pour une nouvelle). */
function saveCompet(key, f, ref) {
  staff_(key);
  if (!f.d1 || !f.nom) throw new Error('Date et nom obligatoires');
  var old = ref && ref.id ? compParId_(ref.id) : null, v = old ? old.v.slice() : new Array(34).join('.').split('.').map(function () { return ''; });
  function s(x) { return String(x == null ? '' : x).trim(); }
  v[0] = frIso_(f.d1); v[1] = frIso_(f.d2 || f.d1); v[2] = 'Fleuret'; v[3] = f.hf || 'H'; v[4] = f.cat || ''; v[5] = f.ie || 'Indiv'; v[6] = s(f.nom); v[7] = s(f.lieu); v[8] = f.type || ''; v[9] = f.cal || ''; v[10] = frIso_(f.dep); v[11] = frIso_(f.ret);
  if (f.log) {
    v[12] = s(f.chd); v[13] = s(f.ent); v[14] = s(f.arb); v[15] = s(f.deleg); v[16] = s(f.hotel); v[33] = s(f.note);
    (f.tr || []).forEach(function (t, k) { if (k < 4) { v[17 + 4 * k] = frIso_(t.d); v[18 + 4 * k] = s(t.v); v[19 + 4 * k] = s(t.p); v[20 + 4 * k] = s(t.n); } });
  }
  var sh = compSh_(), id = old ? old.id : 'C' + v[0].split('/').reverse().join('') + '-' + Utilities.getUuid().slice(0, 4);
  var lock = LockService.getScriptLock(), warn = ''; lock.waitLock(20000);
  try {
    var row = [id].concat(v).concat([new Date()]);
    if (old) sh.getRange(old.ligne, 1, 1, 36).setValues([row]); else sh.appendRow(row);
    _comps = null;
    try { compVersMaitre_(old ? refMaitre_(old) : null, v); } catch (e) { warn = String(e.message || e); Logger.log('Compétition → maître : ' + e); }
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('cal_n'); cacheClear_();
  return warn ? 'ok (fichier maître non mis à jour : ' + warn + ')' : 'ok';
}
function compVersMaitre_(ref, v) {
  var errs = [], d = function (x) { var p = String(x || '').split('/'); return p.length === 3 ? new Date(+p[2], +p[1] - 1, +p[0]) : ''; };
  var row = [d(v[0]), d(v[1])].concat(v.slice(2, 10)).concat([d(v[10]), d(v[11])]);
  ['Liste compet', 'Priorisation A-D'].forEach(function (tab) {
    var sh = sheet_(IDS.MAITRE, tab), l = ref ? ligneCompet_(sh, ref) : 0;
    if (!l) {
      var vv = sh.getRange(4, 1, Math.max(sh.getLastRow() - 3, 1), 1).getValues(), pos = sh.getLastRow() + 1;
      for (var i = 0; i < vv.length; i++) if (vv[i][0] instanceof Date && vv[i][0] > row[0]) { pos = i + 4; break; }
      if (pos <= sh.getLastRow()) sh.insertRowBefore(pos);
      l = pos;
    }
    // onglets en « Tableau » à colonnes typées : on n'écrit que les cases qui changent, une par une
    var cur = sh.getRange(l, 1, 1, 12).getDisplayValues()[0];
    row.forEach(function (val, c) { if (String(cur[c]) !== String(v[c])) try { sh.getRange(l, c + 1).setValue(val); } catch (e) { errs.push(tab + ' col ' + (c + 1)); } });
    if (tab === 'Liste compet') {
      var cl = sh.getRange(l, 13, 1, 22).getDisplayValues()[0];
      v.slice(12, 34).forEach(function (x, i) { if (String(cl[i]) !== String(x)) try { sh.getRange(l, 13 + i).setValue([5, 9, 13, 17].indexOf(i) >= 0 && x ? d(x) : x); } catch (e) { errs.push(tab + ' col ' + (13 + i)); } });
    }
  });
  if (errs.length) throw new Error('cases refusées : ' + errs.join(', '));
}
function supprimerCompet(key, ref) {
  staff_(key);
  var c = compParId_(ref.id), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    compSh_().deleteRow(c.ligne); _comps = null;
    var sp = sheet_(IDS.BASE, 'Priorités');
    if (sp.getLastRow() > 1) { var v = sp.getRange(2, 1, sp.getLastRow() - 1, 1).getDisplayValues(); for (var i = v.length - 1; i >= 0; i--) if (v[i][0] === c.id) sp.deleteRow(i + 2); }
    try { ['Liste compet', 'Priorisation A-D'].forEach(function (tab) { var sh = sheet_(IDS.MAITRE, tab), l = ligneCompet_(sh, refMaitre_(c)); if (l) sh.deleteRow(l); }); } catch (e) { Logger.log('Suppression maître : ' + e); }
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('cal_n'); cacheClear_();
  return 'ok';
}
/* Bandeaux de la semaine (départs, compétitions, retours) et encadré « compétitions de la semaine », depuis BASE. */
function abrComp_(x) {
  var n = String(x[6]).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, ''), a;
  if (/COUPE DU MONDE/.test(n)) a = 'CDM'; else if (/COUPE EURO/.test(n)) a = 'CEE'; else if (/SATELLITE/.test(n)) a = 'SAT'; else if (/EUROFENCE/.test(n)) a = 'EFL';
  else if (/\bEN ?(\d)/.test(n)) a = 'EN' + n.match(/\bEN ?(\d)/)[1]; else if (/CHAMPIONNATS? D.EUROPE/.test(n)) a = 'CHE'; else if (/CHAMPIONNATS? DU MONDE/.test(n)) a = 'CHM';
  else if (/CHAMPIONNAT DE FRANCE/.test(n)) a = 'CF' + (/ELITE/.test(n) ? ' ÉLITE' : ''); else if (/\bGP\b|GRAND PRIX/.test(n)) a = 'GP'; else a = String(x[6]).trim();
  return String(x[4]).trim() + ' - ' + a;
}
function compSemaine_(dates) {
  var lun = dates[0], dim = dates[6], d7 = datesSem_(fmt_(new Date(new Date(lun + 'T12:00:00').getTime() + 7 * 864e5), 'yyyy-MM-dd'));
  var cs = comps_().filter(function (c) { return String(c.v[3]).indexOf('H') >= 0 && !/pas de participation/i.test(c.v[8]) && isoFr_(c.v[0]); });
  var evt = {}, compets = [], suivantes = [];
  function add(d, t) { if (d >= lun && d <= dim) (evt[d] = evt[d] || []).push(t); }
  function per(x) { var a = String(x[0]).slice(0, 5), b = String(x[1]).slice(0, 5); return a === b || !b ? a : a + '–' + b; }
  cs.forEach(function (c) {
    var x = c.v, d1 = isoFr_(x[0]), d2 = isoFr_(x[1]) || d1, dep = isoFr_(x[10]), ret = isoFr_(x[11]), ab = abrComp_(x), lieu = String(x[7]).toUpperCase();
    if (dep && dep < d1) add(dep, '🛫 DÉPART • ' + ab + ' - ' + lieu);
    for (var d = new Date(d1 + 'T12:00:00'); fmt_(d, 'yyyy-MM-dd') <= d2; d.setDate(d.getDate() + 1)) add(fmt_(d, 'yyyy-MM-dd'), '🤺 COMPET • ' + ab + ' - ' + lieu);
    if (ret && ret > d2) add(ret, '🛬 RETOUR • ' + ab + ' - ' + lieu);
    if (d1 <= dim && d2 >= lun) compets.push({ date: per(x), nom: ab, lieu: lieu, depart: String(x[10]).slice(0, 5), retour: String(x[11]).slice(0, 5) });
    if (d1 >= d7[0] && d1 <= d7[6]) suivantes.push(per(x) + ' • ' + ab + ' - ' + lieu);
  });
  return { evt: evt, compets: compets, suivantes: suivantes };
}
function appliqueComp_(s, dates) {
  try {
    var cp = compSemaine_(dates);
    s.compets = cp.compets; s.suivantes = cp.suivantes;
    s.jours.forEach(function (J) {
      var garde = String(J.evt || '').split('\n').filter(function (l) { return l.trim() && !/^(🛫|🤺|🛬)/.test(l.trim()); });
      var uniq = []; (cp.evt[J.date] || []).forEach(function (l) { if (uniq.indexOf(l) < 0) uniq.push(l); });
      J.evt = garde.concat(uniq).join('\n');
    });
  } catch (e) { Logger.log('appliqueComp_ : ' + e); }
  return s;
}
/* Planning annuel imprimable : contenu des entraînements collectifs jour par jour sur toute la saison (PP / F / A). */
function getEntrainementsSaison(key, forcer) {
  staff_(key);
  var C = CacheService.getScriptCache(), c = !forcer && C.get('ent_saison');
  if (c) return JSON.parse(c);
  var deb = new Date(DEBUT_SAISON + 'T12:00:00'), out = {};
  for (var w = 0; w < 53; w++) {
    var l = new Date(deb); l.setDate(l.getDate() + 7 * w);
    var off = Math.round((l - lundi_(0)) / 6048e5), sm = null;
    try { sm = semaine_(off); } catch (e) { }
    if (!sm || sm.provisoire) continue;
    sm.jours.forEach(function (J) {
      var k = [];
      J.creneaux.forEach(function (cr) { var q = String(cr.quoi).toUpperCase(); if (/PR[ÉE]PA|\bPP\b|MUSCU/.test(q) && k.indexOf('PP') < 0) k.push('PP'); if (/FONDA/.test(q) && k.indexOf('F') < 0) k.push('F'); if (/ASSAUT/.test(q) && k.indexOf('A') < 0) k.push('A'); });
      if (k.length) out[J.date] = k.join(' ');
    });
  }
  try { C.put('ent_saison', JSON.stringify(out), 21600); } catch (e) { }
  return out;
}
/* Import unique depuis le fichier maître. */
function importerCompets() {
  var ss = ss_(IDS.BASE);
  if (ss.getSheetByName('Compétitions')) throw new Error('Déjà importé');
  var lc = sheet_(IDS.MAITRE, 'Liste compet'), v = lc.getRange(1, 1, lc.getLastRow(), 34).getDisplayValues();
  var head = ['ID'].concat(v[2].map(function (h) { return String(h).replace(/\s*\n\s*/g, ' ').trim(); })).concat(['Mis à jour']);
  var sh = ss.insertSheet('Compétitions'); sh.getRange(1, 1, sh.getMaxRows(), 36).setNumberFormat('@');
  var rows = [], vus = {};
  for (var i = 3; i < v.length; i++) {
    if (!v[i][0] || !v[i][6]) continue;
    var b = 'C' + v[i][0].split('/').reverse().join(''), id = b, k = 1; while (vus[id]) id = b + '-' + (++k); vus[id] = 1;
    rows.push([id].concat(v[i]).concat(['']));
  }
  sh.getRange(1, 1, 1, 36).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1);
  sh.getRange(2, 1, rows.length, 36).setValues(rows);
  // colonnes du calendrier et priorités
  var pz = sheet_(IDS.MAITRE, 'Priorisation A-D'), pv = pz.getRange(1, 1, pz.getLastRow(), pz.getLastColumn()).getDisplayValues();
  var ct = ss.insertSheet('Calendrier tireurs'), tir = profilsLecons_().filter(function (p) { return p.role === 'Tireur'; });
  var libs = [], g = '', g0 = '';
  for (var c = 12; c < pv[2].length; c++) {
    if (pv[0][c]) { g0 = String(pv[0][c]).trim(); g = ''; } if (pv[1][c]) g = pv[1][c];
    var lib = String(pv[2][c]).trim(); if (!lib) continue;
    var w = lib.split(/\s+/);
    libs.push({ lib: lib, pre: norm_(w[0]), rest: norm_(w.slice(1).join(' ')).replace(/ /g, ''), grp: (/partenaire/i.test(g0) ? 'Partenaires · ' : '') + (g || g0), nom: '' });
  }
  var pris = {};
  libs.slice().sort(function (a, b) { return b.rest.length - a.rest.length; }).forEach(function (L) {
    var c1 = tir.filter(function (p) {
      if (pris[p.nom]) return false;
      var pn = norm_(p.prenom), n = norm_(p.nom).replace(/ /g, ''), ini = norm_(p.nom).split(' ').map(function (x) { return x.charAt(0); }).join('');
      return (pn.indexOf(L.pre) === 0 || L.pre.indexOf(pn) === 0) && L.rest && (n.indexOf(L.rest) === 0 || ini.indexOf(L.rest) === 0);
    });
    if (c1.length) { L.nom = c1[0].nom; pris[c1[0].nom] = 1; }
  });
  var ctr = libs.map(function (L, i) { return [i + 1, L.lib, L.nom, L.grp]; });
  ct.getRange(1, 1, 1, 4).setValues([['Ordre', 'Libellé', 'Tireur (nom court)', 'Groupe']]).setFontWeight('bold'); ct.setFrozenRows(1);
  ct.getRange(2, 1, ctr.length, 4).setValues(ctr);
  var pr = ss.insertSheet('Priorités'), prr = [];
  pr.getRange(1, 1, 1, 4).setValues([['ID compétition', 'Libellé tireur', 'Priorité', 'Mis à jour']]).setFontWeight('bold'); pr.setFrozenRows(1);
  var key = function (a) { return a[0] + '|' + String(a[6]).trim() + '|' + a[7] + '|' + a[3]; }, ids = {};
  rows.forEach(function (r) { ids[key(r.slice(1))] = ids[key(r.slice(1))] || r[0]; });
  for (var i2 = 3; i2 < pv.length; i2++) {
    var id2 = ids[key(pv[i2])]; if (!id2) continue;
    for (var c2 = 12; c2 < pv[2].length; c2++) if (pv[2][c2] && String(pv[i2][c2]).trim()) prr.push([id2, String(pv[2][c2]).trim(), String(pv[i2][c2]).trim(), '']);
  }
  if (prr.length) pr.getRange(2, 1, prr.length, 4).setValues(prr);
  pr.getRange('A:A').setNumberFormat('@');
  CacheService.getScriptCache().remove('cal_n'); cacheClear_();
  Logger.log(rows.length + ' compétitions, ' + ctr.length + ' colonnes tireurs (' + ctr.map(function (x) { return x[1] + '→' + x[2]; }).join(', ') + '), ' + prr.length + ' priorités');
}
/* Stages, vacances scolaires, jours fériés, regroupements : onglet « Événements » de BASE. */
function evenements_() {
  return rows_(IDS.BASE, 'Événements').slice(1).filter(function (x) { return x[0] && x[1]; }).map(function (x) {
    var d1 = isoTxt_(x[1]); return { id: x[0], d1: d1, d2: isoTxt_(x[2]) || d1, type: x[3], titre: x[4], lieu: x[5], qui: x[6], notes: x[7] };
  });
}
function saveEvenement(key, f) {
  staff_(key);
  if (!f.d1 || !f.titre) throw new Error('Date et titre obligatoires');
  var sh = sheet_(IDS.BASE, 'Événements'), id = f.id || ('E' + new Date().getTime());
  var row = [id, f.d1, f.d2 || f.d1, f.type || 'Stage', f.titre, f.lieu || '', f.qui || '', f.notes || ''];
  var l = trouveLigne_(sh, 1, id);
  if (l) sh.getRange(l, 1, 1, 8).setValues([row]); else sh.appendRow(row);
  CacheService.getScriptCache().remove('cal_n');
  return id;
}
function supprimerEvenement(key, id) {
  staff_(key);
  var sh = sheet_(IDS.BASE, 'Événements'), l = trouveLigne_(sh, 1, id);
  if (l) sh.deleteRow(l);
  CacheService.getScriptCache().remove('cal_n');
  return 'ok';
}
/* Écritures du portail dans le MAÎTRE : priorités (Priorisation A-D) et compétitions (Liste compet + Priorisation A-D). */
function ligneCompet_(sh, ref) {
  var v = sh.getRange(1, 1, sh.getLastRow(), 12).getDisplayValues();
  var d = String(ref.d1).split('-'), dt = d[2] + '/' + d[1] + '/' + d[0];
  for (var i = 3; i < v.length; i++) {
    if (v[i][0] === dt && String(v[i][6]).trim() === String(ref.nom).trim() && String(v[i][7]) === String(ref.lieu) && String(v[i][3]).indexOf('H') >= 0) return i + 1;
  }
  return 0;
}
/* =========================================================
 * Disponibilités et planification des leçons (BASE)
 * ======================================================= */
function profilsLecons_() {
  return rows_(IDS.BASE, 'Profils leçons').slice(1).filter(function (x) { return x[0] && String(x[11]).toUpperCase() !== 'NON'; }).map(function (x) {
    return { nom: String(x[0]).trim(), prenom: x[1], role: /coach/i.test(x[2]) ? 'Coach' : 'Tireur', groupe: x[3], cible: num_(x[4]), max: num_(x[5]), maxJour: num_(x[6]), pref1: x[7], pref2: x[8], autorises: x[9], acces: x[10], dbl: num_(x[12]) };
  });
}
function saveObjectifs(key, nom, lecons, doublettes) {
  staff_(key);
  var sh = sheet_(IDS.BASE, 'Profils leçons'), l = trouveLigne_(sh, 1, nom);
  if (!l) throw new Error('Tireur introuvable dans Profils leçons');
  sh.getRange(l, 5).setValue(Number(lecons) || 0); sh.getRange(l, 13).setValue(Number(doublettes) || 0);
  return 'ok';
}
/* Règles de répartition des leçons : priorité des maîtres d'armes selon le rang de la leçon dans la semaine, par groupe. */
function reglesSh_() { return shCree_('Règles leçons', ['Groupe', 'Leçon n°', '1er choix', '2e choix', '3e choix']); }
function reglesLecons_() {
  var sh = reglesSh_(), o = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).getDisplayValues().forEach(function (x) {
    if (!x[0] || !x[1]) return; var g = norm_(x[0]) || 'TOUS';
    (o[g] = o[g] || {})[Number(x[1])] = x.slice(2).map(function (c) { return String(c).trim(); }).filter(Boolean);
  });
  return o;
}
/* r = { objectifs: { POLISTE: { l: 2, d: 1 }, ... }, prio: { POLISTE: { 1: ['GUY DL', 'TOM V'], ... }, ... } } */
function saveReglesLecons(key, r) {
  staff_(key);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (r.objectifs) {
      var sh = sheet_(IDS.BASE, 'Profils leçons'), n = sh.getLastRow() - 1, v = sh.getRange(2, 1, n, 13).getValues(), ch = 0;
      v.forEach(function (x) {
        if (!x[0] || /coach/i.test(x[2])) return;
        var o = r.objectifs[norm_(x[3])]; if (!o) return;
        if (o.l !== '' && o.l != null) { x[4] = Number(o.l) || 0; ch++; }
        if (o.d !== '' && o.d != null) x[12] = Number(o.d) || 0;
      });
      if (ch) { sh.getRange(2, 5, n, 1).setValues(v.map(function (x) { return [x[4]]; })); sh.getRange(2, 13, n, 1).setValues(v.map(function (x) { return [x[12]]; })); }
    }
    if (r.prio) {
      var rs = reglesSh_(), rows = [];
      Object.keys(r.prio).forEach(function (g) { Object.keys(r.prio[g]).forEach(function (k) { var L = (r.prio[g][k] || []).filter(Boolean); if (L.length) rows.push([g, Number(k), L[0] || '', L[1] || '', L[2] || '']); }); });
      if (rs.getLastRow() > 1) rs.getRange(2, 1, rs.getLastRow() - 1, 5).clearContent();
      if (rows.length) rs.getRange(2, 1, rows.length, 5).setValues(rows);
    }
  } finally { lock.releaseLock(); }
  return { profils: profilsLecons_(), regles: reglesLecons_() };
}
function dispos_(lundi) {
  var out = { type: {}, sem: {} };
  rows_(IDS.BASE, 'Dispos').slice(1).forEach(function (x) {
    if (!x[1]) return;
    var s = String(x[0]).trim().toUpperCase() === 'TYPE' ? 'TYPE' : isoTxt_(x[0]);
    if (s !== 'TYPE' && s !== lundi) return;
    var p = []; for (var j = 0; j < 6; j++) p.push([x[3 + 4 * j], x[4 + 4 * j], x[5 + 4 * j], x[6 + 4 * j]]);
    (s === 'TYPE' ? out.type : out.sem)[String(x[1]).trim()] = { p: p, note: x[27], maj: x[28] };
  });
  return out;
}
function monProfil_(u, P0) {
  var n = norm_(u.nom), P = P0 || profilsLecons_();
  var a = P.filter(function (p) { return p.acces && norm_(p.acces) === n; })[0]; if (a) return a;
  if (u.role !== 'Tireur') return null;
  /* Nom d'abord, puis prénom souple (orthographe différente entre Accès et Profils leçons, ex. Zackariya / Zakariya). */
  var c = P.filter(function (p) { return !p.acces && p.role === 'Tireur' && (' ' + n).indexOf(' ' + norm_(p.nom)) >= 0; });
  if (c.length <= 1) return c[0] || null;
  var ex = c.filter(function (p) { return !p.prenom || n.indexOf(norm_(p.prenom)) >= 0; }); if (ex.length) return ex[0];
  var pr = c.filter(function (p) { return n.split(' ').some(function (w) { return w.slice(0, 3) === norm_(p.prenom).slice(0, 3); }); });
  return pr[0] || null;
}
function getMesDispos(key, lundi) {
  var u = user_(key); if (!u) throw new Error('Lien invalide');
  var p = monProfil_(u);
  if (!p) throw new Error('Ton nom n\'est pas encore dans l\'onglet « Profils leçons » : demande à Guy de t\'ajouter.');
  var d = dispos_(lundi);
  return { nom: p.nom, role: p.role, lundi: lundi, sem: d.sem[p.nom] || null, type: d.type[p.nom] || null };
}
function getDisposEquipe(key, lundi) {
  var u = staff_(key), d = dispos_(lundi), moi = monProfil_(u);
  return { lundi: lundi, sem: d.sem, type: d.type, profils: profilsLecons_().map(function (p) {
    return { nom: p.nom, role: p.role, rempli: !!d.sem[p.nom], type: !!d.type[p.nom], moi: !!(moi && moi.nom === p.nom) };
  }) };
}
function saveDispos(key, nom, lundi, plages, note) {
  var u = user_(key); if (!u) throw new Error('Lien invalide');
  if (u.role !== 'Staff') { var p = monProfil_(u); if (!p || p.nom !== nom) throw new Error('Tu ne peux modifier que tes propres disponibilités'); }
  evt_('dispos', u, u.role === 'Staff' && u.nom !== nom ? '' : u.nom);
  var prof = profilsLecons_().filter(function (x) { return x.nom === nom; })[0];
  var sem = lundi || 'TYPE', sh = sheet_(IDS.BASE, 'Dispos');
  var row = [sem, nom, prof ? prof.role : ''];
  for (var j = 0; j < 6; j++) for (var k = 0; k < 4; k++) row.push((plages[j] && plages[j][k]) || '');
  row.push(note || '', new Date());
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues() : [], l = 0;
    for (var i = 0; i < v.length; i++) { var s = String(v[i][0]).trim().toUpperCase() === 'TYPE' ? 'TYPE' : isoTxt_(v[i][0]); if (s === sem && String(v[i][1]).trim() === nom) { l = i + 2; break; } }
    if (!l) l = sh.getLastRow() + 1;
    sh.getRange(l, 1, 1, row.length).setValues([row]);
  } finally { lock.releaseLock(); }
  return 'ok';
}
function leconsPlanifiees_(lundi) {
  var sh = sheet_(IDS.BASE, 'Leçons planifiées');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 11).getDisplayValues().filter(function (x) { return x[0] && isoTxt_(x[1]) === lundi; })
    .map(function (x) { return { id: x[0], date: isoTxt_(x[2]), h: x[3], duree: num_(x[4]) || 30, coach: x[5], tireur: x[6], vs: x[7], statut: x[8], note: x[9] }; });
}
function getPlanif(key, lundi) {
  staff_(key);
  var d = dispos_(lundi), prof = profilsLecons_(), disp = {};
  prof.forEach(function (p) { var s = d.sem[p.nom], t = d.type[p.nom]; disp[p.nom] = { p: s ? s.p : t ? t.p : [], src: s ? 'semaine' : t ? 'type' : 'aucune', note: s ? s.note : '' }; });
  var off = Math.round((new Date(lundi + 'T12:00:00') - lundi_(0)) / 6048e5), s = null;
  try { s = semaine_(off); } catch (e) { }
  return { lundi: lundi, profils: prof, dispo: disp, lecons: leconsPlanifiees_(lundi), regles: reglesLecons_(),
    jours: s ? s.jours.map(function (j) { return { date: j.date, jour: j.jour, evt: j.evt, creneaux: j.creneaux }; }) : null, maitre: s ? s.jours.map(function (j) { return j.lecons.length; }) : null };
}
function saveLecons(key, lundi, liste, supprimer) {
  var u = staff_(key), out = saveLecons_(lundi, liste, supprimer);
  try {
    var lock = LockService.getScriptLock(); lock.waitLock(30000);
    try {
      if (lundi >= fmt_(lundi_(0), 'yyyy-MM-dd') && (liste || []).some(function (x) { return /valid/i.test(x.statut || 'validée'); })) basculer_(lundi, u.nom);
      if (gerees_()[lundi]) ecrireMaitre_(lundi);
    } finally { lock.releaseLock(); }
    cacheClear_(); CacheService.getScriptCache().remove('saison_n');
  } catch (e) { Logger.log('ecrireMaitre_ : ' + e); }
  return leconsPlanifiees_(lundi);
}
function saveLecons_(lundi, liste, supprimer, sansVerrou) {
  var sh = sheet_(IDS.BASE, 'Leçons planifiées'), lock = sansVerrou ? { releaseLock: function () { } } : LockService.getScriptLock(); if (!sansVerrou) lock.waitLock(20000);
  try {
    var ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (x) { return x[0]; }) : [];
    supprimer = supprimer || [];
    if (supprimer.length > 3) { // suppression en bloc : on réécrit le tableau sans les lignes supprimées (bien plus rapide que ligne par ligne)
      var sup = {}; supprimer.forEach(function (id) { sup[id] = 1; });
      var all = sh.getRange(2, 1, ids.length, 11).getValues(), sem = sh.getRange(2, 2, ids.length, 1).getDisplayValues(), garde = all.filter(function (x, i) { return !sup[ids[i]] || isoTxt_(sem[i][0]) !== lundi; });
      if (garde.length !== all.length) { sh.getRange(2, 1, all.length, 11).clearContent(); if (garde.length) sh.getRange(2, 1, garde.length, 11).setValues(garde); }
      ids = garde.map(function (x) { return String(x[0]); });
    } else supprimer.forEach(function (id) { var i = ids.lastIndexOf(id); if (i >= 0) { sh.deleteRow(i + 2); ids.splice(i, 1); } });
    var nouvelles = [], seq = 0, t0 = new Date().getTime(), deja = {};
    if (ids.length) sh.getRange(2, 1, ids.length, 7).getDisplayValues().forEach(function (x) { deja[[isoTxt_(x[2]), x[3], x[5], x[6]].join('|')] = 1; });
    (liste || []).forEach(function (f) {
      if (!f.id) { var kd = [f.date, f.h, f.coach, f.tireur].join('|'); if (deja[kd]) return; deja[kd] = 1; } // double clic : pas de doublon
      var id = f.id || ('L' + t0 + ('00' + (seq++)).slice(-3) + Math.floor(Math.random() * 10)); f.id = id;
      var row = [id, lundi, f.date, f.h, f.duree || 30, f.coach, f.tireur, f.vs || '', f.statut || 'validée', f.note || '', new Date()];
      var i = ids.indexOf(id);
      if (i >= 0) sh.getRange(i + 2, 1, 1, 11).setValues([row]); else nouvelles.push(row);
    });
    if (nouvelles.length) sh.getRange(sh.getLastRow() + 1, 1, nouvelles.length, 11).setValues(nouvelles);
  } finally { lock.releaseLock(); }
  return leconsPlanifiees_(lundi);
}
/* Rappel hebdomadaire « donne tes disponibilités » (non installé tant que le portail n'est pas lancé). */
function envoyerDemandeDispos() {
  var url = ScriptApp.getService().getUrl(), lun = fmt_(lundi_(1), 'yyyy-MM-dd'), d = dispos_(lun);
  acces_().forEach(function (a) {
    if (!a.rappel || !a.email || ['Tireur', 'Coach', 'Staff'].indexOf(a.role) < 0) return;
    var p = monProfil_(a); if (p && d.sem[p.nom]) return; // déjà rempli pour la semaine prochaine
    MailApp.sendEmail({ to: a.email, subject: 'Tes disponibilités pour la semaine du ' + lun.split('-').reverse().join('/'),
      htmlBody: 'Bonjour,<br><br>Indique tes créneaux disponibles de la semaine prochaine (2 plages par jour) : <a href="' + url + '?k=' + encodeURIComponent(a.cle) + '#dispos">ouvrir mes disponibilités</a>.<br><br>Pôle France Relève fleuret H' });
  });
}

function getPresences(key) {
  staff_(key);
  var sh = sheet_(IDS.BASE, 'Appel'), out = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(function (x) {
    if (!x[2]) return;
    var o = out[x[2]] = out[x[2]] || { tireur: x[2], total: 0, retardMin: 0 };
    o.total++; o[x[3]] = (o[x[3]] || 0) + 1; o.retardMin += Number(x[4]) || 0;
  });
  return Object.keys(out).map(function (k) { return out[k]; });
}
function getParcoursTireur(key, nom) {
  if (!user_(key)) throw new Error('Lien invalide');
  var c = CacheService.getScriptCache(), ck = 'pt_' + norm_(nom).replace(/ /g, '_'), v = c.get(ck);
  if (v) return JSON.parse(v);
  var p = parcours_(nom).map(function (x) { return { ep: x.ep, date: x.date, place: x.place, txt: x.txt }; });
  try { c.put(ck, JSON.stringify(p), 21600); } catch (e) { }
  return p;
}

/* =========================================================
 * Tâches planifiées
 * ======================================================= */

function envoyerRappelsWellness() {
  /* Notification seulement (plus d'e-mail), chaque matin sauf le dimanche, aux tireurs du pôle qui n'ont pas encore répondu. */
  if (Number(fmt_(new Date(), 'u')) === 7) return;
  var pr = PropertiesService.getScriptProperties(), td0 = fmt_(new Date(), 'yyyy-MM-dd'); if (pr.getProperty('WELL_ENVOYE') === td0) return; pr.setProperty('WELL_ENVOYE', td0);
  var td = td0, deja = {}; wellness_(1).forEach(function (w) { if (w.d === td) deja[norm_(w.tireur)] = 1; });
  var noms = acces_().filter(function (a) { if (a.role !== 'Tireur') return false; var t = trouveTireur_(a.nom); return t && t.pole && !deja[t.cle] && !deja[norm_(a.nom)]; }).map(function (a) { return a.nom; });
  return pushEnvoyer_(noms, 'Wellness du jour', 'Comment tu te sens ce matin ? 30 secondes pour répondre.', '#wellness');
}

/** À lancer une fois : rappel wellness 9h + cache des résultats chaque nuit à 6h45. */
function installerDeclencheurs() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['envoyerRappelsWellness', 'construireCacheResultats'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('envoyerRappelsWellness').timeBased().everyDays(1).atHour(9).nearMinute(0).inTimezone(TZ).create();
  ScriptApp.newTrigger('construireCacheResultats').timeBased().everyDays(1).atHour(6).nearMinute(45).inTimezone(TZ).create();
}

/** Rappel hebdomadaire des disponibilités : chaque mercredi vers 20 h (uniquement aux personnes cochées « Rappel » dans Accès, et pas à ceux qui ont déjà rempli). */
function installerRappelDispos() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'envoyerDemandeDispos') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('envoyerDemandeDispos').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(20).nearMinute(0).inTimezone(TZ).create();
  return ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
}

/** Test : vérifie que toutes les sources se lisent. */
function testerSources() {
  var cle = acces_().filter(function (a) { return a.role === 'Staff'; })[0].cle;
  cacheClear_();
  var t0 = Date.now(), d = getData(cle);
  Logger.log('getData ' + (Date.now() - t0) + ' ms · semaine ' + (d.semaine && d.semaine.onglet) + ' · leçons lundi ' + (d.semaine ? d.semaine.jours[0].lecons.length : '-') + ' · tireurs ' + d.tireurs.length);
  t0 = Date.now(); construireCacheResultats(); Logger.log('cache résultats ' + (Date.now() - t0) + ' ms');
  t0 = Date.now(); var r = getResultats(cle, 'M20 FH'); Logger.log('getResultats ' + (Date.now() - t0) + ' ms · ' + r.lignes.length + ' lignes · ' + r.groupes.length + ' groupes');
  t0 = Date.now(); getData(cle); Logger.log('getData (cache) ' + (Date.now() - t0) + ' ms');
}

/* =========================================================
 * Relais parcours Engarde : le navigateur lit Engarde (refusé à Apps Script),
 * puis envoie ici les poules et tableaux ; le portail met en forme et écrit les notes du V6.
 * ======================================================= */
function doPost(e) {
  var out;
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.action === 'parcours') out = majParcours_(p);
    else if (p.action === 'notesSrc') out = notesDepuisSrc_(p);
    else if (p.action === 'rappelDispos') { staff_(p.k); out = installerRappelDispos(); }
    else if (p.action === 'voix') return ContentService.createTextOutput(voixTexte_(p.k, p.texte)).setMimeType(ContentService.MimeType.TEXT);
    else if (p.action === 'api') { try { out = { ok: api_(p) }; } catch (e2) { out = { erreur: e2.message }; } return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON); }
    else throw new Error('action inconnue');
  } catch (err) {
    out = { erreur: err.message };
    try { journalParcours_({ epreuve: '?', source: '' }, 'ERREUR ' + err.message); } catch (e2) {}
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function journalParcours_(p, statut) {
  var sh = sheet_(IDS.RESULTATS, 'SOURCES_PARCOURS');
  sh.appendRow([p.onglets || '', p.tag || '', p.epreuve, p.source, 'ENGARDE via navigateur', 'OUI', statut, new Date()]);
}

function majParcours_(p) {
  staff_(p.k);
  var sx = p.sexe, cat = p.cat;
  var onglets = ['M17 ' + sx, 'M20 ' + sx, 'M23 ' + sx, 'Senior ' + sx];
  // rangs nationaux actuels
  var rk = {};
  [['M17', 'M17 ' + sx], ['M20', 'M20 ' + sx], ['Senior', 'Senior ' + sx]].forEach(function (c) {
    var sh = sheet_(IDS.RESULTATS, c[1]), l = sh.getLastRow(); if (l < 5) return;
    sh.getRange(5, 3, l - 4, 6).getValues().forEach(function (r) {
      if (!r[0] || r[5] === '' || r[5] == null) return;
      var k = norm_(r[0] + ' ' + r[1]); (rk[k] = rk[k] || {})[c[0]] = r[5];
    });
  });
  var ordre = cat === 'Senior' ? ['Senior', 'M20'] : ['M20', 'M17'];
  function lab(nom) {
    if (/\([A-Z]{3}\)$/.test(String(nom))) return '';
    var r = rk[norm_(nom)]; if (!r) return 'non classé';
    var parts = ordre.filter(function (c) { return r[c]; }).map(function (c) { return (c === 'Senior' ? 'Sr' : c) + ' ' + r[c]; });
    if (!parts.length && r.Senior) parts = ['Sr ' + r.Senior];
    return parts.length ? 'CN ' + parts.join(' · ') : 'non classé';
  }
  function placeTxt(n) { n = Number(n); return n ? (n === 1 ? '1er' : n + 'e') : ''; }
  var d = p.date.split('-'), dd = d[2] + '-' + d[1], dateFr = d[2] + '/' + d[1] + '/' + d[0];
  // repérage des colonnes et des lignes
  var ville = norm_(p.ville), CAT = norm_(cat);
  var cibles = [];
  onglets.forEach(function (nom) {
    var sh = sheet_(IDS.RESULTATS, nom); if (!sh) return;
    var lc = sh.getLastColumn(), lr = sh.getLastRow(); if (lr < 5) return;
    var h = sh.getRange(1, 1, 3, lc).getDisplayValues();
    var noms = sh.getRange(5, 3, lr - 4, 2).getValues().map(function (r) { return r[0] ? norm_(r[0] + ' ' + r[1]) : ''; });
    for (var c = 0; c < lc; c++) {
      var tag = String(h[0][c]), hd = String(h[2][c]).replace(/\n/g, ' '), H = norm_(hd);
      if (!/^(RESULTATS|CHRONO|SELECTIF)/.test(tag) || /PTS|TOTAL|EUROPE|MONDE/.test(H)) continue;
      if (H.indexOf(ville) < 0) continue;
      var mh = hd.match(/(\d{1,2})-(\d{1,2})-(\d{2,4})/);
      if (mh) { var yh = mh[3].length === 2 ? '20' + mh[3] : mh[3], dh = new Date(yh + '-' + ('0' + mh[2]).slice(-2) + '-' + ('0' + mh[1]).slice(-2) + 'T12:00:00'); if (Math.abs(dh - new Date(p.date + 'T12:00:00')) > 5 * 864e5) continue; }
      else { var sai = String(tag.split('|')[2] || ''), an = Number(d[0]), mo = Number(d[1]), sp = mo >= 8 ? an + '/' + (an + 1) : (an - 1) + '/' + an; if (sai && sai !== sp) continue; }
      var okCat = /^RESULTATS/.test(tag) ? String(tag.split('|')[1]).toUpperCase() === String(cat).toUpperCase() : (' ' + hd.toUpperCase() + ' ').indexOf(' ' + String(cat).toUpperCase() + ' ') >= 0;
      if (!okCat) continue;
      cibles.push({ sh: sh, onglet: nom, col: c + 1, res: /^RESULTATS/.test(tag), noms: noms });
    }
  });
  // places déjà dans le fichier (colonne résultats)
  var placeFichier = {};
  cibles.filter(function (t) { return t.res; }).forEach(function (t) {
    t.sh.getRange(5, t.col, t.noms.length, 1).getValues().forEach(function (v, i) { if (t.noms[i] && v[0] !== '' && placeFichier[t.noms[i]] == null) placeFichier[t.noms[i]] = v[0]; });
  });
  // mise en forme
  var notes = {}, places = {}, src = [];
  (p.tireurs || []).forEach(function (f) {
    var k = norm_(f.nom), pl = p.placesEngarde ? f.place : (placeFichier[k] || f.place);
    var L = [p.epreuve + ' · ' + dateFr];
    if (pl) L.push(placeTxt(pl) + (p.nb ? ' sur ' + p.nb : ''));
    var multi = (f.poules || []).some(function (x) { return x.tour > 1; });
    (f.poules || []).forEach(function (x, i) {
      var diff = x.TS - x.TR, last = i === f.poules.length - 1;
      L.push('');
      L.push((x.phase ? x.phase + ' · ' : '') + (multi ? 'TOUR ' + x.tour + ' · ' : '') + 'POULE ' + x.no + ' · ' + x.V + ' V, ' + x.D + ' D · ' + x.TS + '/' + x.TR + ' (' + (diff >= 0 ? '+' : '') + diff + ')' + (x.rg ? ' · ' + placeTxt(x.rg) + ' après poules' : ''));
      x.b.forEach(function (b) { L.push('  ' + b[0] + ' ' + b[1] + '-' + b[2] + '  ' + b[3] + (lab(b[3]) ? ' (' + lab(b[3]) + ')' : '')); });
    });
    if ((f.tab || []).length) {
      L.push(''); L.push('TABLEAU');
      f.tab.forEach(function (b) { L.push('  ' + (b[5] ? b[5] + ' ' : '') + b[0] + '  ' + b[1] + ' ' + b[2] + '-' + b[3] + '  ' + b[4] + (lab(b[4]) ? ' (' + lab(b[4]) + ')' : '')); });
    } else if ((f.poules || []).length) { L.push(''); L.push('Éliminé à l\'issue des poules'); }
    L.push(''); L.push('CN = classement national FFE au ' + (p.cnDate || '')); L.push('Source : ' + (p.srcLib || 'Engarde (' + p.source + ')'));
    notes[k] = L.join('\n'); if (f.place) places[k] = f.place;
    src.push([p.epreuve, new Date(p.date + 'T12:00:00'), p.cle, f.nom, pl ? placeTxt(pl) + (p.nb ? ' sur ' + p.nb : '') : '', notes[k]]);
  });
  // écriture des notes (et des places Engarde si demandé)
  var nNotes = 0, nPlaces = 0, ong = {};
  cibles.forEach(function (t) {
    var rg = t.sh.getRange(5, t.col, t.noms.length, 1), nt = rg.getNotes(), ch = false;
    t.noms.forEach(function (k, i) { if (k && notes[k]) { nt[i][0] = notes[k]; ch = true; nNotes++; } });
    if (ch) { rg.setNotes(nt); ong[t.onglet] = 1; }
    if (t.res && p.placesEngarde) {
      var v = rg.getValues(), chv = false;
      t.noms.forEach(function (k, i) { if (k && places[k] && v[i][0] !== places[k]) { v[i][0] = places[k]; chv = true; nPlaces++; } });
      if (chv) rg.setValues(v);
    }
  });
  // archive SRC_PARCOURS (remplace la même épreuve)
  var sp = sheet_(IDS.RESULTATS, 'SRC_PARCOURS'), lr = sp.getLastRow();
  var old = lr > 1 ? sp.getRange(2, 1, lr - 1, 6).getValues().filter(function (r) { return r[0] && String(r[2]) !== String(p.cle); }) : [];
  var all = old.concat(src);
  if (lr > 1) sp.getRange(2, 1, lr - 1, 6).clearContent();
  if (all.length + 1 > sp.getMaxRows()) sp.insertRowsAfter(sp.getMaxRows(), all.length + 1 - sp.getMaxRows());
  if (all.length) sp.getRange(2, 1, all.length, 6).setValues(all);
  var bilan = { notes: nNotes, places: nPlaces, colonnes: cibles.map(function (t) { return t.onglet + ' col ' + t.col; }), tireurs: (p.tireurs || []).length };
  p.onglets = Object.keys(ong).join(', ');
  journalParcours_(p, 'OK ' + nNotes + ' notes · ' + nPlaces + ' places corrigées · ' + bilan.colonnes.length + ' colonnes');
  return bilan;
}

/* Rattrapage : pose en note les parcours déjà archivés (SRC_PARCOURS) sur les cases de résultat qui n'en ont pas. */
function notesDepuisSrc_(p) {
  staff_(p.k);
  var sp = sheet_(IDS.RESULTATS, 'SRC_PARCOURS'), lr = sp.getLastRow(), idx = {};
  if (lr > 1) sp.getRange(2, 1, lr - 1, 6).getValues().forEach(function (r) {
    if (!r[0] || !r[5]) return;
    var E = norm_(r[0]), U = String(r[0]).toUpperCase(), cat = /\bM17\b/.test(U) ? 'M17' : /\bM20\b/.test(U) ? 'M20' : 'SENIOR';
    var d = r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'dd-MM') : String(r[1]).replace(/^(\d\d)\/(\d\d).*$/, '$1-$2');
    var ville = norm_(String(r[0]).replace(/^(EI|EN|CN|TI|CF)\s+(\d+\s+)?/i, '').replace(/\s(M17|M20|Senior|N\d)\b.*$/i, ''));
    var k = cat + '|' + d; var x = idx[k] = idx[k] || { villes: {}, n: {} };
    x.villes[ville] = 1; x.n[norm_(r[3])] = r[5];
  });
  var total = 0, detail = [];
  ['FH', 'FD'].forEach(function (sx) {
    ['M17', 'M20', 'M23', 'Senior'].forEach(function (c) {
      var nom = c + ' ' + sx, sh = sheet_(IDS.RESULTATS, nom); if (!sh) return;
      var lc = sh.getLastColumn(), lr2 = sh.getLastRow(); if (lr2 < 5) return;
      var h = sh.getRange(1, 1, 3, lc).getDisplayValues();
      var noms = sh.getRange(5, 3, lr2 - 4, 2).getValues().map(function (r) { return r[0] ? norm_(r[0] + ' ' + r[1]) : ''; });
      for (var j = 0; j < lc; j++) {
        var tag = String(h[0][j]), hd = String(h[2][j]).replace(/\n/g, ' '), H = norm_(hd);
        if (!/^(RESULTATS|CHRONO|SELECTIF)/.test(tag) || /PTS|TOTAL|EUROPE|MONDE/.test(H)) continue;
        var m = hd.match(/(\d\d)-(\d\d)-\d{2,4}/); if (!m) continue;
        var HU = hd.toUpperCase(), cat = /^RESULTATS/.test(tag) ? String(tag.split('|')[1]).toUpperCase() : (/\bM17\b/.test(HU) ? 'M17' : /\bM20\b/.test(HU) ? 'M20' : 'SENIOR');
        var x = idx[cat + '|' + m[1] + '-' + m[2]]; if (!x) continue;
        if (!Object.keys(x.villes).some(function (v) { return v && H.indexOf(v) >= 0; })) continue;
        var rg = sh.getRange(5, j + 1, noms.length, 1), v = rg.getValues(), nt = rg.getNotes(), ch = 0;
        noms.forEach(function (k, i) { if (k && v[i][0] !== '' && v[i][0] != null && !nt[i][0] && x.n[k]) { nt[i][0] = x.n[k]; ch++; } });
        if (ch) { rg.setNotes(nt); total += ch; detail.push(nom + ' ' + hd + ' : ' + ch); }
      }
    });
  });
  journalParcours_({ epreuve: 'Rattrapage notes depuis SRC_PARCOURS', source: '' }, 'OK ' + total + ' notes ajoutées · ' + detail.join(' | '));
  return { notes: total, detail: detail };
}

/* =========================================================
 * Sélection (Résultats) : deux listes ordonnées par catégorie
 * Onglet SELECTION de BASE : Catégorie | Liste | Rang | Tireur | Mis à jour
 * ======================================================= */
function selSheet_() {
  var ss = ss_(IDS.BASE), sh = ss.getSheetByName('SELECTION');
  if (!sh) { sh = ss.insertSheet('SELECTION'); sh.getRange(1, 1, 1, 5).setValues([['Catégorie', 'Liste', 'Rang', 'Tireur', 'Mis à jour']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function getSelection(key, cat) {
  var u = selOk_(key, cat), P = panel_(cat);
  var out = { panel: P, listes: {}, moi: u.nom, peut: P.some(function (p) { return p.nom === u.nom; }) };
  var sh = selSheet_(), n = sh.getLastRow();
  if (n < 2) return out;
  var v = sh.getRange(2, 1, n - 1, 4).getValues().filter(function (r) { return r[0] === cat; });
  v.sort(function (a, b) { return a[2] - b[2]; });
  v.forEach(function (r) { (out.listes[r[1]] = out.listes[r[1]] || []).push(String(r[3])); });
  return out;
}
function saveSelection(key, cat, noms) {
  var u = selOk_(key, cat);
  if (!panel_(cat).some(function (p) { return p.nom === u.nom; })) throw new Error('Tu ne fais pas partie du groupe de sélection de cette catégorie');
  var liste = u.nom;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = selSheet_(), n = sh.getLastRow(), garde = [];
    if (n > 1) garde = sh.getRange(2, 1, n - 1, 5).getValues().filter(function (r) { return r[0] && !(r[0] === cat && r[1] === liste); });
    var now = new Date();
    (noms || []).forEach(function (nm, i) { garde.push([cat, liste, i + 1, nm, now]); });
    if (n > 1) sh.getRange(2, 1, n - 1, 5).clearContent();
    if (garde.length) sh.getRange(2, 1, garde.length, 5).setValues(garde);
  } finally { lock.releaseLock(); }
  return 'ok';
}


/* =========================================================
 * Saisie vocale des scores (raccourci Siri)
 * « Mahel gagne 15-3 contre Smaïn », « Despois perd 4-5 contre Chino puis Anane bat Crist 5-2 »
 * ======================================================= */
var VX_NB = { ZERO: 0, UN: 1, UNE: 1, DEUX: 2, TROIS: 3, QUATRE: 4, CINQ: 5, SIX: 6, SEPT: 7, HUIT: 8, NEUF: 9, DIX: 10, ONZE: 11, DOUZE: 12, TREIZE: 13, QUATORZE: 14, QUINZE: 15 };
var VX_STOP = {}; 'GAGNE GAGNER BAT BATTU PERD PERDU CONTRE FACE LE LA LES L A AU ET DE DU DES SUR PAR MATCH MATCHS TOUCHE TOUCHES POULE TABLEAU SCORE VICTOIRE DEFAITE EMPORTE REMPORTE INCLINE S SE CONTRE ENSUITE PUIS APRES ALORS IL ELLE QUI'.split(' ').forEach(function (w) { VX_STOP[w] = 1; });
function vxCle_(w) { // clé « phonétique » simple : sans accents, sans H, Y=I, lettres doublées réduites
  return norm_(w).replace(/ /g, '').replace(/PH/g, 'F').replace(/H/g, '').replace(/Y/g, 'I').replace(/(.)\1+/g, '$1').replace(/E?S$/, '');
}
function vxDist_(a, b) {
  if (Math.abs(a.length - b.length) > 2) return 9;
  var d = []; for (var i = 0; i <= a.length; i++) { d[i] = [i]; for (var j = 1; j <= b.length; j++) d[i][j] = i ? Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)) : j; }
  return d[a.length][b.length];
}
function vxTireurs_() {
  var l = tireurs_().filter(function (t) { return t.nom && t.prenom && !/ancien|arr[êe]t/i.test(t.statut); });
  var cpt = {}; l.forEach(function (t) { var k = vxCle_(t.prenom); cpt[k] = (cpt[k] || 0) + 1; });
  return l.map(function (t) {
    return { t: t, nom: t.nom + ' ' + t.prenom, kp: vxCle_(t.prenom), kn: vxCle_(String(t.nom).split(/[\s-]+/)[0]), kn2: vxCle_(t.nom), ini: norm_(t.nom).charAt(0), ambigu: cpt[vxCle_(t.prenom)] > 1, pole: t.pole };
  });
}
/* Repère les tireurs dans une liste de mots ; renvoie [{c, pos}] dans l'ordre. */
function vxNoms_(mots, cands) {
  var trouves = [], pris = {};
  for (var i = 0; i < mots.length; i++) {
    if (VX_STOP[norm_(mots[i])]) continue;
    var w = vxCle_(mots[i]), w2 = i + 1 < mots.length && !VX_STOP[norm_(mots[i + 1])] ? vxCle_(mots[i] + mots[i + 1]) : null;
    if (w.length < 3 && !w2) continue;
    var best = null;
    cands.forEach(function (c) {
      [[w, 1], [w2, 2]].forEach(function (x) {
        if (!x[0] || x[0].length < 3) return;
        [['kp', c.kp], ['kn', c.kn], ['kn2', c.kn2]].forEach(function (y) {
          var dd = vxDist_(x[0], y[1]), tol = y[1].length >= 6 ? 2 : y[1].length >= 4 ? 1 : 0;
          if (dd > tol) return;
          var sc = dd * 10 + (x[1] === 2 ? 0 : 1) + (c.pole ? 0 : 3);
          if (y[0] === 'kp' && c.ambigu) { // prénom partagé : on regarde le mot suivant (initiale ou nom)
            var nx = mots[i + x[1]] ? norm_(mots[i + x[1]]) : '';
            if (nx && (nx === c.ini || vxDist_(vxCle_(nx), c.kn) <= 1)) sc -= 5; else sc += 20;
          }
          if (!best || sc < best.sc) best = { c: c, sc: sc, len: x[1] };
        });
      });
    });
    if (best && best.sc < 20 && !pris[best.c.nom]) { trouves.push({ c: best.c, pos: i }); pris[best.c.nom] = 1; i += best.len - 1; }
  }
  return trouves;
}
function vxParse_(texte, seuls) {
  var cands = vxTireurs_(), out = [], rates = [];
  if (seuls && seuls.length) { var ok = {}; seuls.forEach(function (n) { ok[norm_(n)] = 1; }); var cc = cands.filter(function (c) { return ok[norm_(c.nom)]; }); if (cc.length >= 2) cands = cc; }
  String(texte || '').replace(/[’']/g, ' ').split(/\b(?:puis|ensuite|apr[eè]s|match suivant|et puis)\b|[.;\n]/i).forEach(function (seg) {
    seg = seg.trim(); if (!seg) return;
    var brut = seg.replace(/(\d+)\s*(?:-|–|—|à|a|:|\/|contre)\s*(\d+)/gi, ' $1 $2 ');
    var mots = brut.split(/[\s,!?]+/).filter(Boolean), nums = [], motsNoms = [];
    mots.forEach(function (m) { var n = norm_(m); if (/^\d+$/.test(m)) nums.push(Number(m)); else if (VX_NB[n] != null && !(n === 'UN' || n === 'UNE') ) nums.push(VX_NB[n]); else if ((n === 'UN' || n === 'UNE') && nums.length) nums.push(1); else motsNoms.push(m); });
    var noms = vxNoms_(motsNoms, cands);
    if (noms.length < 2 || nums.length < 2) { rates.push(seg); return; }
    var a = noms[0].c, b = noms[1].c, sa = nums[0], sb = nums[1], low = norm_(seg);
    var gagne = /\b(GAGNE|BAT|BATTU|L EMPORTE|REMPORTE|VAINQUEUR|VICTOIRE)\b/.test(low), perd = /\b(PERD|PERDU|DEFAITE|BATTU PAR|S INCLINE)\b/.test(low);
    if (/BATTU PAR/.test(low)) { gagne = false; perd = true; }
    if (gagne && sa < sb || perd && sa > sb) { var z = sa; sa = sb; sb = z; }
    if (sa === sb || sa > 45 || sb > 45) { rates.push(seg); return; }
    out.push({ a: a.nom, b: b.nom, sa: sa, sb: sb, pa: a.t.prenom + (a.ambigu ? ' ' + a.t.nom : ''), pb: b.t.prenom + (b.ambigu ? ' ' + b.t.nom : '') });
  });
  var amb = {}; cands.forEach(function (c) { if (c.ambigu && rates.some(function (x) { return x && vxCle_(x).indexOf(c.kp) >= 0; })) amb[c.t.prenom] = 1; });
  return { matchs: out, rates: rates, ambigus: Object.keys(amb) };
}
/* Séance d'assauts du jour la plus proche de l'heure actuelle ; sinon on en crée une. */
function vxSeance_(key) {
  var now = new Date(), iso = fmt_(now, 'yyyy-MM-dd'), h = Number(fmt_(now, 'H')) + Number(fmt_(now, 'm')) / 60;
  var sh = sheet_(IDS.BASE, 'Séances'), v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
  var best = null;
  v.forEach(function (x) {
    var d = x[1] instanceof Date ? fmt_(x[1], 'yyyy-MM-dd') : String(x[1]);
    if (!x[0] || d !== iso || !/assaut|match|combat|tir/i.test(String(x[4]))) return;
    var m = String(x[2]).match(/(\d{1,2})\s*h\s*(\d{0,2})/i), hh = m ? Number(m[1]) + (Number(m[2]) || 0) / 60 : 12, ec = Math.abs(hh - h);
    if (!best || ec < best.ec) best = { id: String(x[0]), h: String(x[2]), type: String(x[4]), ec: ec };
  });
  if (best) return best;
  var hTxt = fmt_(now, 'H') + 'h';
  var id = saveSeance(key, { date: iso, h: hTxt, type: 'Assauts', notes: 'Créée par la saisie vocale' });
  return { id: id, h: hTxt, type: 'Assauts', cree: true };
}
function vxExos_(idS) {
  var sh = sheet_(IDS.BASE, 'Exercices'); if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 10).getDisplayValues().filter(function (x) { return x[0] === String(idS) && x[9]; }).map(function (x) {
    var p = String(x[9]).split(/\s*,\s*/).filter(String);
    return { idE: x[1], format: x[3], consignes: x[8], p: p, k: p.map(norm_) };
  });
}
function voixTexte_(key, texte) {
  try {
    var u = staff_(key);
    texte = String(texte || '').trim();
    if (!texte) return "Je n'ai rien entendu.";
    if (/^\s*(annule|supprime|efface)/i.test(texte)) return vxAnnule_(u);
    var r = vxParse_(texte);
    var aide = r.ambigus.length ? ' Il y a plusieurs ' + r.ambigus.join(', ') + ' : ajoute le nom.' : '';
    var lowM = ' ' + norm_(texte.replace(/[’']/g, ' ')) + ' ', estMemo = MEMO_RX.some(function (x) { return x[1].test(lowM); });
    if (!r.matchs.length || (estMemo && !/\b(GAGNE|BAT|PERD|CONTRE)\b/.test(lowM))) {
      if (!r.matchs.length && !estMemo && /\d+\s*(?:-|à|a|:|\/|\s)\s*\d+/.test(texte)) return "Je n'ai pas compris le score « " + texte + " »." + (aide || ' Dis par exemple : Mahel gagne 15 à 3 contre Smaïn.');
      var mm = memoEnregistrer_(u, memoAnalyse_(texte));
      return memoPhrase_(mm);
    }
    var se = vxSeance_(key), sh = sheet_(IDS.BASE, 'Assauts'), d = new Date(fmt_(new Date(), 'yyyy-MM-dd') + 'T12:00:00');
    var mo = moduleEtat_(se.id);
    if (mo && (mo.tours || []).length) {
      var tousM = []; mo.tours.forEach(function (t) { mxGroupes_(t).forEach(function (g) { g.tireurs.forEach(function (n) { if (tousM.indexOf(n) < 0) tousM.push(n); }); }); });
      var r3 = vxParse_(texte, tousM); if (r3.matchs.length >= r.matchs.length) r = r3;
      var pl = vxModule_(u, se, r.matchs);
      if (pl) {
        var rest = r.matchs.filter(function (m) { return !m.place; });
        return 'Noté : ' + r.matchs.filter(function (m) { return m.place; }).map(function (m) { return m.pa + ' ' + m.sa + ', ' + m.pb + ' ' + m.sb; }).join(' ; ') + ' (' + pl.join(' ; ') + ').' +
          (rest.length ? ' Pas trouvé dans les poules : ' + rest.map(function (m) { return m.pa + ' contre ' + m.pb; }).join(', ') + '.' : '') + (r.rates.length ? ' Pas compris : ' + r.rates.join(' ; ') + '.' : '');
      }
    }
    // Poules / exercices de la séance : on reconnaît les noms parmi les participants et on range chaque match dans sa poule.
    var exs = vxExos_(se.id), tous = [];
    exs.forEach(function (e) { e.p.forEach(function (n) { if (tous.indexOf(n) < 0) tous.push(n); }); });
    if (tous.length) { var r2 = vxParse_(texte, tous); if (r2.matchs.length >= r.matchs.length) r = r2; }
    r.matchs.forEach(function (m) {
      var ka = norm_(m.a), kb = norm_(m.b);
      m.ex = exs.filter(function (e) { return e.k.indexOf(ka) >= 0 && e.k.indexOf(kb) >= 0; }).pop() || null; // la poule la plus récente (2e tour, etc.)
    });
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var rows = r.matchs.map(function (m) { return [d, se.type, m.a, m.b, m.sa, m.sb, 'voix', m.ex ? m.ex.format : Math.max(m.sa, m.sb) > 5 ? '15 touches' : '5 touches', u.nom, se.type, se.id, m.ex ? m.ex.idE : '']; });
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
    } finally { lock.releaseLock(); }
    cacheClear_();
    var dit = r.matchs.map(function (m) { var po = m.ex && String(m.ex.consignes).match(/^Poule \d+/); return (po ? po[0] + ' : ' : '') + m.pa + ' ' + m.sa + ', ' + m.pb + ' ' + m.sb; }).join(' ; ');
    return 'Noté' + (r.matchs.length > 1 ? ' ' + r.matchs.length + ' matchs' : '') + ' : ' + dit + '. Séance ' + se.type + ' de ' + se.h + (se.cree ? ', créée' : '') + '.' +
      (r.rates.length ? ' Pas compris : ' + r.rates.join(' ; ') + '.' + aide : '');
  } catch (e) { return 'Erreur : ' + e.message; }
}
function vxAnnule_(u) {
  var sh = sheet_(IDS.BASE, 'Assauts'), l = sh.getLastRow();
  if (l < 2) return 'Rien à annuler.';
  var v = sh.getRange(Math.max(2, l - 30), 1, Math.min(30, l - 1) + 1, 12).getValues(), base = Math.max(2, l - 30);
  for (var i = v.length - 1; i >= 0; i--) {
    if (v[i][6] === 'voix' && v[i][8] === u.nom) {
      var m = v[i]; sh.deleteRow(base + i); cacheClear_();
      return 'Annulé : ' + m[2] + ' ' + m[4] + ', ' + m[3] + ' ' + m[5] + '.';
    }
  }
  return 'Aucun score dicté à annuler.';
}

/* =========================================================
 * Module « Assauts » (logiciel d'escrime intégré à la séance)
 * L'état complet d'une séance (tours, poules, tableaux, scores, abandons, retards) est gardé en JSON
 * dans l'onglet « Rencontres » de BASE, puis recopié en lignes dans Exercices, Assauts et Appel.
 * ======================================================= */
function rencontresSheet_() {
  var ss = SpreadsheetApp.openById(IDS.BASE), sh = ss.getSheetByName('Rencontres');
  if (!sh) { sh = ss.insertSheet('Rencontres'); sh.getRange(1, 1, 1, 5).setValues([['ID séance', 'Date', 'État du module (JSON)', 'Mis à jour', 'Par']]); sh.setFrozenRows(1); }
  return sh;
}
function moduleEtat_(idS) {
  var sh = rencontresSheet_(), l = trouveLigne_(sh, 1, idS);
  if (!l) return null;
  try { return JSON.parse(sh.getRange(l, 3).getValue() || 'null'); } catch (e) { return null; }
}
function getModule(key, idS) { staff_(key); return moduleEtat_(idS); }
function dicterScore(key, texte) { return voixTexte_(key, texte); }
/* Réécrit un onglet en retirant les lignes ciblées et en ajoutant les nouvelles (une seule écriture). */
function remplacerLignes_(sh, nc, cible, nouvelles) {
  var last = sh.getLastRow(), v = last > 1 ? sh.getRange(2, 1, last - 1, nc).getValues() : [];
  var garde = v.filter(function (r) { return !cible(r); });
  if (garde.length === v.length) { if (nouvelles.length) sh.getRange(last + 1, 1, nouvelles.length, nc).setValues(nouvelles); return; }
  if (v.length) sh.getRange(2, 1, v.length, nc).clearContent();
  var tout = garde.concat(nouvelles);
  if (tout.length) sh.getRange(2, 1, tout.length, nc).setValues(tout);
}
function mxCle_(a, b) { return a < b ? a + '|' + b : b + '|' + a; }
function mxGroupes_(t) { // [{idx, nom, tireurs}]
  if (t.type === 'poules') return (t.groupes || []).map(function (g, i) { return { idx: i + 1, nom: 'Poule ' + (i + 1) + '/' + t.groupes.length, tireurs: g }; });
  return [{ idx: 1, nom: t.type === 'tableau' ? 'Tableau' : t.type === 'equipes' ? 'Équipes' : 'Matchs', tireurs: t.tireurs || [] }];
}
function moduleLignes_(idS, date, st, u) {
  var d = new Date(date + 'T12:00:00'), ex = [], ma = [], ordre = 0;
  (st.tours || []).forEach(function (t, ti) {
    var ab = t.ab || {}, fmt = { poules: 'Poules', tableau: 'Tableau', libre: 'Matchs libres', equipes: 'Matchs par équipe', roi: 'Roi de la piste' }[t.type] || 'Matchs libres';
    var lib = (t.type === 'poules' ? 'Poule ' : t.type === 'tableau' ? 'Tableau ' : 'Match ') + (t.touches || 15) + ' touches';
    if (t.type === 'equipes') {
      ordre++;
      var res = Object.keys(t.sc || {}).map(function (k) { var m = t.sc[k]; return m.a + ' ' + m.sa + '-' + m.sb + ' ' + m.b; });
      ex.push([idS, idS + '-T' + (ti + 1) + 'G1', ordre, fmt, t.touches || '', '', t.intensite || '', '', '[Assauts] ' + (t.nom || 'Tour ' + (ti + 1)) + ' · ' + (t.equipes || []).map(function (e) { return e.nom + ' : ' + e.m.join(', '); }).join(' | ') + (res.length ? ' · Résultats : ' + res.join(' ; ') : ''), (t.tireurs || []).join(', '), '', res.length]);
      return;
    }
    mxGroupes_(t).forEach(function (g) {
      var idE = idS + '-T' + (ti + 1) + 'G' + g.idx, n = 0, mem = {};
      g.tireurs.forEach(function (x) { mem[x] = 1; });
      Object.keys(t.sc || {}).forEach(function (k) {
        var m = t.sc[k];
        if (!mem[m.a] || !mem[m.b] || ab[m.a] === 'annuler' || ab[m.b] === 'annuler') return;
        if (m.sa === '' || m.sb === '' || m.sa == null || m.sb == null) return;
        n++;
        ma.push([d, 'Assauts', m.a, m.b, Number(m.sa), Number(m.sb), 'module', lib, u.nom, t.nom || ('Tour ' + (ti + 1)), idS, idE]);
      });
      ordre++;
      var t5 = Number(t.touches) <= 5;
      ex.push([idS, idE, ordre, fmt, t.touches || 15, '', t.intensite || '', '', '[Assauts] ' + (t.nom || 'Tour ' + (ti + 1)) + ' · ' + g.nom, g.tireurs.join(', '), t5 ? n : '', t5 ? '' : n]);
    });
  });
  return { ex: ex, ma: ma };
}
function moduleEnregistrer_(u, idS, date, st) {
  var json = JSON.stringify(st);
  if (json.length > 48000) throw new Error('Séance trop volumineuse pour être enregistrée.');
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var sh = rencontresSheet_(), l = trouveLigne_(sh, 1, idS), row = [idS, new Date(date + 'T12:00:00'), json, new Date(), u.nom];
    if (l) sh.getRange(l, 1, 1, 5).setValues([row]); else sh.appendRow(row);
    var L = moduleLignes_(idS, date, st, u), pre = idS + '-T';
    var shE = sheet_(IDS.BASE, 'Exercices');
    remplacerLignes_(shE, 12, function (r) { return String(r[0]) === idS && String(r[1]).indexOf(pre) === 0; }, L.ex);
    var shA = sheet_(IDS.BASE, 'Assauts');
    remplacerLignes_(shA, 12, function (r) { return String(r[10]) === idS && String(r[11]).indexOf(pre) === 0; }, L.ma);
    // Appel : tout tireur du module est présent ; abandon → Écourté ; retard → Retard (sans écraser une saisie plus précise)
    var shP = sheet_(IDS.BASE, 'Appel'), last = shP.getLastRow(), v = last > 1 ? shP.getRange(2, 1, last - 1, 9).getValues() : [];
    var mien = v.filter(function (r) { return String(r[0]) === idS; }), parNom = {};
    mien.forEach(function (r) { parNom[r[2]] = r; });
    var d = new Date(date + 'T12:00:00'), now = new Date(), voulu = {};
    (st.tours || []).forEach(function (t) {
      mxGroupes_(t).forEach(function (g) { g.tireurs.forEach(function (n) { if (!voulu[n]) voulu[n] = 'Présent'; }); });
      Object.keys(t.retard || {}).forEach(function (n) { if (voulu[n] !== 'Écourté') voulu[n] = 'Retard'; });
      Object.keys(t.ab || {}).forEach(function (n) { if (t.ab[n]) voulu[n] = 'Écourté'; });
    });
    var change = false;
    Object.keys(voulu).forEach(function (n) {
      var r = parNom[n], s = voulu[n];
      if (!r) { parNom[n] = [idS, d, n, s, '', '', s === 'Écourté' ? 'Abandon en cours de séance' : '', now, u.nom]; change = true; return; }
      var cur = String(r[3] || '');
      if (s === 'Présent' && /^Absent|^$/.test(cur)) { r[3] = 'Présent'; change = true; }
      else if (s === 'Retard' && /^Présent$|^Absent|^$/.test(cur)) { r[3] = 'Retard'; change = true; }
      else if (s === 'Écourté' && !/Écourté|Adapté/.test(cur)) { r[3] = 'Écourté'; if (!r[6]) r[6] = 'Abandon en cours de séance'; change = true; }
    });
    if (change) remplacerLignes_(shP, 9, function (r) { return String(r[0]) === idS; }, Object.keys(parNom).map(function (n) { return parNom[n]; }));
  } finally { lock.releaseLock(); }
  cacheClear_();
  return { ok: true, maj: fmt_(new Date(), 'HH:mm:ss') };
}
function saveModule(key, idS, date, st) { var u = staff_(key); return moduleEnregistrer_(u, idS, date, st); }
/* Saisie vocale : si la séance a un module, le score va dans le tour le plus récent qui contient les deux tireurs. */
function vxModule_(u, se, matchs) {
  var st = moduleEtat_(se.id); if (!st || !(st.tours || []).length) return null;
  var date = fmt_(new Date(), 'yyyy-MM-dd'), place = [];
  matchs.forEach(function (m) {
    for (var i = st.tours.length - 1; i >= 0; i--) {
      var t = st.tours[i]; if (t.type === 'equipes' || t.type === 'roi') continue;
      var g = mxGroupes_(t).filter(function (g) { return g.tireurs.indexOf(m.a) >= 0 && g.tireurs.indexOf(m.b) >= 0; })[0];
      if (!g) continue;
      t.sc = t.sc || {};
      t.sc[mxCle_(m.a, m.b)] = { a: m.a, b: m.b, sa: m.sa, sb: m.sb };
      place.push((t.nom || 'Tour ' + (i + 1)) + (t.type === 'poules' ? ', ' + g.nom.replace(/\/\d+$/, '') : ''));
      m.place = true; return;
    }
  });
  if (!place.length) return null;
  moduleEnregistrer_(u, se.id, date, st);
  return place;
}

/* =========================================================
 * Mémos vocaux : score, absence, rendez-vous médical, scolaire, leçon, tâche, note
 * Onglet « Mémos » de BASE : Créé le | Par | Type | Tireur(s) | Pour le | Heure | Texte | Statut | Fait le | ID
 * ======================================================= */
var MEMO_LIB = { absence: 'Absence', medical: 'Médical', scolaire: 'Scolaire', lecon: 'Leçon', tache: 'Tâche', note: 'Note' };
var MEMO_RX = [
  ['absence', /\b(ABSENT|ABSENTE|ABSENTS|ABSENCE|NE VIENT PAS|NE VIENDRA PAS|VIENDRA PAS|VIENT PAS|PAS LA|PAS PRESENT|MALADE|EXCUSE|EXCUSEE|DISPENSE|DISPENSEE|EN RETARD|ARRIVERA EN RETARD|PARTIRA PLUS TOT|PART PLUS TOT)\b/],
  ['medical', /\b(MEDECIN|DOCTEUR|DOC|KINE|OSTEO|OSTEOPATHE|PODOLOGUE|DENTISTE|ORTHODONTISTE|RADIO|IRM|ECHO|ECHOGRAPHIE|SCANNER|BLESSE|BLESSEE|BLESSURE|DOULEUR|DOULEURS|ENTORSE|TENDINITE|MAL AU|MAL A LA|MAL AUX|SOINS|INFIRMERIE|CERTIFICAT|CONSULTATION|URGENCES|ANTIDOPAGE)\b/],
  ['scolaire', /\b(EXAMEN|EXAMENS|CONTROLE|CONTROLES|BAC|BREVET|PARTIEL|PARTIELS|DS|DST|DEVOIR|DEVOIRS|BULLETIN|CONSEIL DE CLASSE|PROF|PROFESSEUR|LYCEE|COLLEGE|FAC|ORAL|INTERRO|EPREUVE ANTICIPEE|STAGE D OBSERVATION|NOTES)\b/],
  ['lecon', /\b(LECON|LECONS)\b/],
  ['tache', /\b(PENSER|PENSE A|RAPPELER|RAPPELLE MOI|RAPPEL|APPELER|APPELLE|ENVOYER|ENVOIE|COMMANDER|COMMANDE|ACHETER|PREVOIR|PREPARER|RESERVER|RELANCER|IL FAUT|A FAIRE|N OUBLIE PAS|NE PAS OUBLIER|NOUBLIE PAS|MAIL|IMPRIMER|INSCRIRE|SIGNER|VERIFIER|DEMANDER|VOIR AVEC|TODO)\b/]
];
function memosSheet_() {
  var ss = ss_(IDS.BASE), sh = ss.getSheetByName('Mémos');
  if (!sh) { sh = ss.insertSheet('Mémos'); sh.getRange(1, 1, 1, 10).setValues([['Créé le', 'Par', 'Type', 'Tireur(s)', 'Pour le', 'Heure', 'Texte', 'Statut', 'Fait le', 'ID']]); sh.setFrozenRows(1); }
  return sh;
}
function memoDate_(low) {
  var now = new Date(), base = new Date(fmt_(now, 'yyyy-MM-dd') + 'T12:00:00'), d = null;
  function plus(n) { return new Date(base.getTime() + n * 864e5); }
  if (/\bAPRES DEMAIN\b/.test(low)) d = plus(2);
  else if (/\bDEMAIN\b/.test(low)) d = plus(1);
  else if (/\b(AUJOURD HUI|CE SOIR|CE MATIN|CET APRES MIDI|TOUT A L HEURE|CE MIDI)\b/.test(low)) d = plus(0);
  else if (/\bSEMAINE PROCHAINE\b/.test(low)) { var w = base.getDay(); d = plus(((8 - w) % 7) || 7); }
  if (!d) {
    var J = ['DIMANCHE', 'LUNDI', 'MARDI', 'MERCREDI', 'JEUDI', 'VENDREDI', 'SAMEDI'];
    for (var i = 0; i < 7; i++) if (new RegExp('\\b' + J[i] + '\\b').test(low)) { var e = (i - base.getDay() + 7) % 7; d = plus(e === 0 ? 7 : e); if (new RegExp('\\b' + J[i] + ' PROCHAIN').test(low) && e < 7 && e > 0) d = plus(e); break; }
  }
  if (!d) {
    var MO = ['JANVIER', 'FEVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 'JUILLET', 'AOUT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DECEMBRE'];
    var m = low.match(/\b(\d{1,2})\s*\/\s*(\d{1,2})\b/), j = null, mo = null;
    if (m) { j = Number(m[1]); mo = Number(m[2]) - 1; }
    else { m = low.match(/\b(?:LE\s+)?(\d{1,2}|PREMIER|1ER)\s+(JANVIER|FEVRIER|MARS|AVRIL|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE)\b/); if (m) { j = /^\d/.test(m[1]) ? Number(m[1]) : 1; mo = MO.indexOf(m[2]); } }
    if (j == null) { m = low.match(/\bLE\s+(\d{1,2})\b(?!\s*H)/); if (m) { j = Number(m[1]); mo = base.getMonth(); if (j < base.getDate()) mo++; } }
    if (j != null && j >= 1 && j <= 31) { var y = base.getFullYear(); if (mo < base.getMonth() - 1) y++; d = new Date(y, mo, j, 12); }
  }
  var h = low.match(/\b(\d{1,2})\s*H\s*(\d{2})?\b/);
  return { date: d ? fmt_(d, 'yyyy-MM-dd') : '', heure: h ? h[1] + 'h' + (h[2] || '') : '' };
}
/* Analyse d'un mémo : type, tireurs, date. */
function memoAnalyse_(texte) {
  var low = ' ' + norm_(String(texte).replace(/[’']/g, ' ')) + ' ', type = '';
  for (var i = 0; i < MEMO_RX.length; i++) if (MEMO_RX[i][1].test(low)) { type = MEMO_RX[i][0]; break; }
  var mots = String(texte).replace(/[’']/g, ' ').split(/[\s,.;!?]+/).filter(Boolean);
  var noms = vxNoms_(mots.filter(function (m) { return !/^\d/.test(m); }), vxTireurs_()).map(function (x) { return x.c.nom; });
  var dt = memoDate_(low);
  if (!type) type = noms.length ? 'note' : 'tache';
  return { type: type, tireurs: noms, date: dt.date, heure: dt.heure, texte: String(texte).trim() };
}
var MEMO_DERNIER = null;
function memoEnregistrer_(u, m) {
  var sh = memosSheet_(), id = 'M' + Date.now().toString(36) + Math.floor(Math.random() * 100);
  sh.appendRow([new Date(), u.nom, m.type, (m.tireurs || []).join(', '), m.date ? new Date(m.date + 'T12:00:00') : '', m.heure || '', m.texte, m.type === 'tache' || m.type === 'lecon' ? 'À faire' : '', '', id]);
  m.id = id; MEMO_DERNIER = m;
  return m;
}
function memoPhrase_(m) {
  var d = m.date ? fmt_(new Date(m.date + 'T12:00:00'), 'EEEE dd/MM').replace(/^(\w+)/, function (x) { return ({ Monday: 'lundi', Tuesday: 'mardi', Wednesday: 'mercredi', Thursday: 'jeudi', Friday: 'vendredi', Saturday: 'samedi', Sunday: 'dimanche' })[x] || x; }) : '';
  return 'Mémo noté : ' + MEMO_LIB[m.type] + (m.tireurs.length ? ' · ' + m.tireurs.map(function (n) { var t = trouveTireur_(n); return t ? t.prenom : n; }).join(', ') : '') + (d ? ' · ' + d : '') + (m.heure ? ' ' + m.heure : '') + '.';
}
function memos_() {
  var sh = memosSheet_(), last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 10).getValues().filter(function (x) { return x[9]; }).map(function (x) {
    return { cree: x[0] instanceof Date ? fmt_(x[0], 'yyyy-MM-dd HH:mm') : String(x[0]), par: x[1], type: x[2], tireurs: String(x[3]).split(/\s*,\s*/).filter(String), date: x[4] instanceof Date ? fmt_(x[4], 'yyyy-MM-dd') : String(x[4] || ''), heure: x[5], texte: x[6], statut: x[7], fait: x[8] instanceof Date ? fmt_(x[8], 'yyyy-MM-dd') : '', id: String(x[9]) };
  });
}
/* Mémos utiles pour une semaine : tâches non faites + tout ce qui est daté dans la semaine (ou sans date, créé il y a moins de 3 jours). */
function memosSemaine_(lundiIso) {
  var fin = fmt_(new Date(new Date(lundiIso + 'T12:00:00').getTime() + 6 * 864e5), 'yyyy-MM-dd'), lim = fmt_(new Date(Date.now() - 3 * 864e5), 'yyyy-MM-dd');
  return memos_().filter(function (m) { return m.statut === 'À faire' || (m.date && m.date >= lundiIso && m.date <= fin) || (!m.date && m.cree.slice(0, 10) >= lim); });
}
/* Tâche ajoutée depuis la page Aujourd'hui (toujours de type « tâche »). */
function memoTache(key, texte) {
  var u = staff_(key), t = String(texte || '').trim(); if (!t) throw new Error('Tâche vide');
  var a = memoAnalyse_(t); a.type = 'tache';
  return memoEnregistrer_(u, a).id;
}
function memoVocal(key, texte) { MEMO_DERNIER = null; var msg = voixTexte_(key, texte); return { msg: msg, memo: MEMO_DERNIER }; }
function memoMaj(key, id, ch) {
  staff_(key);
  var sh = memosSheet_(), l = trouveLigne_(sh, 10, id); if (!l) throw new Error('Mémo introuvable');
  if (ch.supprimer) { sh.deleteRow(l); return 'supprimé'; }
  if (ch.type) { sh.getRange(l, 3).setValue(ch.type); if ((ch.type === 'tache' || ch.type === 'lecon') && !sh.getRange(l, 8).getValue()) sh.getRange(l, 8).setValue('À faire'); }
  if (ch.tireurs) sh.getRange(l, 4).setValue(ch.tireurs.join(', '));
  if (ch.date != null) sh.getRange(l, 5).setValue(ch.date ? new Date(ch.date + 'T12:00:00') : '');
  if (ch.texte != null) sh.getRange(l, 7).setValue(ch.texte);
  if (ch.statut != null) { sh.getRange(l, 8).setValue(ch.statut); sh.getRange(l, 9).setValue(ch.statut === 'Fait' ? new Date() : ''); }
  return 'ok';
}

/* =========================================================
 * Application (PWA) : API, notifications push, RPE
 * La page installée sur le téléphone appelle le script par doPost {action:'api', fn, args}.
 * Notifications Web Push « sans contenu » (signature VAPID ES256) : le téléphone vient ensuite
 * chercher le message avec mesNotifs(key). Les clés VAPID sont créées par le script lui-même au premier appel
 * et gardées dans ses propriétés (VAPID_D privée, VAPID_PUB publique) : elles ne sortent jamais du script.
 * ======================================================= */
var API_OK = ['mesNonLues', 'marquerLues', 'getMesCompets', 'getUtilisation', 'setPeInsep', 'majPlaces', 'majClassementsSenior', 'getEntrainementsSaison', 'saveReglesLecons', 'supprimerCompet', 'getSemaineType', 'enregistrerSemaineType', 'copierCreneaux', 'getSaison', 'saveCreneau', 'memoTache', 'getConnexions', 'saveSuiviLecon', 'getSuiviLecons', 'supprimerSuiviLecon', 'actualiserResultats', 'connexion', 'definirMotDePasse', 'getCalendrier', 'getData', 'getDisposEquipe', 'getFiche', 'getImpr', 'getMesDispos', 'getModule', 'getParcoursTireur', 'getParticipants', 'getPlanif', 'getPresences', 'getResultats', 'getSelection', 'getSemaine', 'memoMaj', 'memoVocal', 'rafraichir', 'saveAppel', 'saveChampFiche', 'saveCompet', 'saveDispos', 'saveEvenement', 'saveExercice', 'saveLecons', 'saveMatchs', 'saveModule', 'saveObjectifs', 'savePriorite', 'saveSeance', 'saveSelection', 'supprimerEvenement', 'traiterDemande', 'uploadPhoto', 'saveAssaut', 'saveLame', 'saveScolaire', 'saveTextile', 'saveWellness', 'saveTest', 'saveDette', 'regleDette', 'dicterScore', 'pushAbonner', 'pushDesabonner', 'pushTest', 'mesNotifs', 'getMesSeances', 'saveRPE', 'getVapid', 'envoyerPlanningLecons', 'envoyerMonLien', 'getMuscu', 'saveCharge', 'supprimerCharge', 'getMuscuEquipe'];
function api_(p) {
  if (API_OK.indexOf(p.fn) < 0 || typeof globalThis[p.fn] !== 'function') throw new Error('Fonction inconnue : ' + p.fn);
  return globalThis[p.fn].apply(null, p.args || []);
}
function vapidCles_() {
  var pr = PropertiesService.getScriptProperties(), d = pr.getProperty('VAPID_D'), pub = pr.getProperty('VAPID_PUB');
  if (d && pub) return { d: d, pub: pub };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    d = pr.getProperty('VAPID_D'); pub = pr.getProperty('VAPID_PUB'); if (d && pub) return { d: d, pub: pub };
    var C = ec_(), x = BigInt(0);
    while (x === BigInt(0)) x = ecMod_(BigInt('0x' + bytesHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now(), Utilities.Charset.UTF_8))), C.n);
    var Q = ecMul_(x, C.G);
    function pad(v) { var t = v.toString(16); while (t.length < 64) t = '0' + t; return t; }
    d = pad(x); pub = b64u_(hexBytes_('04' + pad(Q[0]) + pad(Q[1])));
    pr.setProperties({ VAPID_D: d, VAPID_PUB: pub });
    return { d: d, pub: pub };
  } finally { lock.releaseLock(); }
}
function getVapid() { return vapidCles_().pub; }

/* ---- ECDSA P-256 (ES256) en JavaScript pur ---- */
var EC_ = null;
function ec_() {
  if (EC_) return EC_;
  var B = BigInt;
  EC_ = { p: B('0xffffffff00000001000000000000000000000000ffffffffffffffffffffffff'), n: B('0xffffffff00000000ffffffffffffffffbce6faada7179e84f3b9cac2fc632551'),
    a: B('0xffffffff00000001000000000000000000000000fffffffffffffffffffffffc'),
    G: [B('0x6b17d1f2e12c4247f8bce6e563a440f277037d812deb33a0f4a13945d898c296'), B('0x4fe342e2fe1a7f9b8ee7eb4a7c0f9e162bce33576b315ececbb6406837bf51f5')] };
  return EC_;
}
function ecMod_(a, m) { var r = a % m; return r < BigInt(0) ? r + m : r; }
function ecInv_(a, m) { var b = ecMod_(a, m), c = m, x0 = BigInt(1), x1 = BigInt(0); while (c !== BigInt(0)) { var q = b / c, t = b - q * c; b = c; c = t; t = x0 - q * x1; x0 = x1; x1 = t; } return ecMod_(x0, m); }
/* coordonnées jacobiennes pour éviter une inversion par étape */
function ecDbl_(P) { var C = ec_(), p = C.p; if (!P || P[2] === BigInt(0)) return P; var X = P[0], Y = P[1], Z = P[2];
  var YY = Y * Y % p, S = BigInt(4) * X * YY % p, ZZ = Z * Z % p, M = (BigInt(3) * X * X + C.a * ZZ * ZZ) % p;
  var X3 = ecMod_(M * M - BigInt(2) * S, p), Y3 = ecMod_(M * (S - X3) - BigInt(8) * YY * YY, p), Z3 = BigInt(2) * Y * Z % p; return [X3, Y3, Z3]; }
function ecAdd_(P, Q) { var p = ec_().p; if (!P) return Q; if (!Q) return P;
  var Z1Z1 = P[2] * P[2] % p, Z2Z2 = Q[2] * Q[2] % p, U1 = P[0] * Z2Z2 % p, U2 = Q[0] * Z1Z1 % p, S1 = P[1] * Q[2] % p * Z2Z2 % p, S2 = Q[1] * P[2] % p * Z1Z1 % p;
  if (U1 === U2) return S1 === S2 ? ecDbl_(P) : null;
  var H = ecMod_(U2 - U1, p), R = ecMod_(S2 - S1, p), HH = H * H % p, HHH = H * HH % p, V = U1 * HH % p;
  var X3 = ecMod_(R * R - HHH - BigInt(2) * V, p), Y3 = ecMod_(R * (V - X3) - S1 * HHH, p), Z3 = P[2] * Q[2] % p * H % p; return [X3, Y3, Z3]; }
function ecMul_(k, G) { var R = null, A = [G[0], G[1], BigInt(1)]; while (k > BigInt(0)) { if (k & BigInt(1)) R = ecAdd_(R, A); A = ecDbl_(A); k >>= BigInt(1); } var p = ec_().p, zi = ecInv_(R[2], p), zi2 = zi * zi % p; return [R[0] * zi2 % p, R[1] * zi2 % p * zi % p]; }
function bytesHex_(b) { return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join(''); }
function hexBytes_(h) { var o = []; for (var i = 0; i < h.length; i += 2) { var v = parseInt(h.substr(i, 2), 16); o.push(v > 127 ? v - 256 : v); } return o; }
function b64u_(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function es256_(texte, dHex) {
  var C = ec_(), n = C.n, d = BigInt('0x' + dHex);
  var h = BigInt('0x' + bytesHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texte, Utilities.Charset.UTF_8)));
  for (var essai = 0; essai < 10; essai++) {
    var g = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, dHex + '|' + texte + '|' + Utilities.getUuid() + essai, Utilities.Charset.UTF_8);
    var k = ecMod_(BigInt('0x' + bytesHex_(g)), n - BigInt(1)) + BigInt(1);
    var r = ecMod_(ecMul_(k, C.G)[0], n); if (r === BigInt(0)) continue;
    var s = ecMod_(ecInv_(k, n) * (h + r * d), n); if (s === BigInt(0)) continue;
    function pad(x) { var t = x.toString(16); while (t.length < 64) t = '0' + t; return t; }
    return b64u_(hexBytes_(pad(r) + pad(s)));
  }
  throw new Error('Signature impossible');
}
function vapidJwt_(aud) {
  var c = CacheService.getScriptCache(), ck = 'vapid_' + aud.replace(/[^a-z0-9]/gi, '_'), v = c.get(ck);
  if (v) return v;
  var dHex = vapidCles_().d;
  var hd = b64u_(Utilities.newBlob(JSON.stringify({ typ: 'JWT', alg: 'ES256' })).getBytes());
  var pl = b64u_(Utilities.newBlob(JSON.stringify({ aud: aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: 'mailto:pole.relève@ffescrime.fr'.replace('è', 'e') })).getBytes());
  var jwt = hd + '.' + pl + '.' + es256_(hd + '.' + pl, dHex);
  c.put(ck, jwt, 6 * 3600);
  return jwt;
}

/* ---- abonnements et file des messages ---- */
/* ---- Charge d'entraînement : onglet « RPE séances » (Date | Tireur | Séance | Durée | RPE | Charge | Commentaire | Saisi par) ---- */
function rpe_(jours) {
  var sh = rpeSheet_(); if (!sh || sh.getLastRow() < 2) return [];
  var lim = new Date(Date.now() - (jours || 35) * 864e5);
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().filter(function (x) { return x[0] instanceof Date && x[0] >= lim && x[1]; }).map(function (x) {
    return { d: fmt_(x[0], 'yyyy-MM-dd'), tireur: String(x[1]), seance: String(x[2] || ''), duree: Number(x[3]) || 0, rpe: Number(x[4]), charge: Number(x[5]) || (Number(x[3]) || 0) * Number(x[4]) };
  });
}
/* ---- Suivi des leçons : ce que chaque maître d'armes note après une leçon ----
 * Onglet « Suivi leçons » de BASE : Date | Maître d'armes | Tireur | Travail | Situations / exercices | Forme (1-5) | Leçon prise (1-5) | Leçon donnée (1-5) | Commentaire | Saisi le | ID */
function suiviSheet_() {
  var ss = ss_(IDS.BASE), sh = ss.getSheetByName('Suivi leçons');
  if (!sh) { sh = ss.insertSheet('Suivi leçons'); sh.getRange(1, 1, 1, 11).setValues([['Date', 'Maître d\'armes', 'Tireur', 'Travail', 'Situations / exercices', 'Forme du tireur (1-5)', 'Qualité leçon prise (1-5)', 'Qualité leçon donnée (1-5)', 'Commentaire', 'Saisi le', 'ID']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function suivi_() {
  var sh = suiviSheet_(), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 11).getValues().filter(function (x) { return x[10] && x[2]; }).map(function (x) {
    return { date: x[0] instanceof Date ? fmt_(x[0], 'yyyy-MM-dd') : String(x[0]), coach: String(x[1]), tireur: String(x[2]), travail: String(x[3]), situations: String(x[4]), forme: Number(x[5]) || '', prise: Number(x[6]) || '', donnee: Number(x[7]) || '', comm: String(x[8]), id: String(x[10]) };
  }).sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
}
function suiviOk_(key) { var u = user_(key); if (!u || (u.role !== 'Staff' && u.role !== 'Coach')) throw new Error('Réservé aux maîtres d\'armes'); return u; }
function saveSuiviLecon(key, f) {
  var u = suiviOk_(key);
  if (!f || !f.tireur) throw new Error('Choisis le tireur');
  if (!String(f.travail || '').trim() && !String(f.situations || '').trim() && !String(f.comm || '').trim()) throw new Error('Dis au moins ce qui a été travaillé');
  if (!f.id) evt_('lecon', u, prenomDe_(u.nom) + ' → ' + prenomDe_(f.tireur));
  function n5(v) { v = Number(v); return v >= 1 && v <= 5 ? v : ''; }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = suiviSheet_(), row = [new Date((f.date || fmt_(new Date(), 'yyyy-MM-dd')) + 'T12:00:00'), u.nom, f.tireur, String(f.travail || '').trim(), String(f.situations || '').trim(), n5(f.forme), n5(f.prise), n5(f.donnee), String(f.comm || '').trim(), new Date()];
    if (f.id) {
      var l = trouveLigne_(sh, 11, f.id); if (!l) throw new Error('Entrée introuvable');
      if (u.role !== 'Staff' && String(sh.getRange(l, 2).getValue()) !== u.nom) throw new Error('Tu ne peux modifier que tes propres notes');
      row[1] = sh.getRange(l, 2).getValue(); sh.getRange(l, 1, 1, 10).setValues([row]); return f.id;
    }
    var id = 'L' + Date.now().toString(36) + Math.floor(Math.random() * 100);
    sh.appendRow(row.concat([id])); return id;
  } finally { lock.releaseLock(); }
}
function supprimerSuiviLecon(key, id) {
  var u = suiviOk_(key), sh = suiviSheet_(), l = trouveLigne_(sh, 11, id); if (!l) return 'ok';
  if (u.role !== 'Staff' && String(sh.getRange(l, 2).getValue()) !== u.nom) throw new Error('Tu ne peux supprimer que tes propres notes');
  sh.deleteRow(l); return 'ok';
}
/* Staff : tout (90 derniers jours) ; maître d'armes : ses propres notes. */
function getSuiviLecons(key) {
  var u = suiviOk_(key), lim = fmt_(new Date(Date.now() - 90 * 864e5), 'yyyy-MM-dd');
  return suivi_().filter(function (x) { return x.date >= lim && (u.role === 'Staff' || x.coach === u.nom); }).reverse();
}
/* ---- Journal des connexions : Accès I (dernière connexion) et J (nombre), visible par Guy seulement ---- */
var ADMIN = ['Guy'];
function vu_(key) {
  try {
    var c = CacheService.getScriptCache(), ck = 'vu_' + String(key).slice(-16);
    if (c.get(ck)) return;
    c.put(ck, '1', 900);
    var a = acces_().filter(function (x) { return x.cle === key; })[0]; if (!a) return;
    var sh = ss_(IDS.BASE).getSheetByName('Accès');
    if (!sh.getRange(1, 9).getValue()) sh.getRange(1, 9, 1, 2).setValues([['Dernière connexion', 'Connexions']]).setFontWeight('bold');
    var r = sh.getRange(a.ligne, 9, 1, 2), v = r.getValues()[0];
    r.setValues([[new Date(), (Number(v[1]) || 0) + 1]]);
    if (!Number(v[1])) evt_('premiere', user_(key));
  } catch (e) {}
}
function getConnexions(key) {
  var u = user_(key); if (!u || ADMIN.indexOf(u.nom) < 0) throw new Error('Réservé à Guy');
  var v = sheet_(IDS.BASE, 'Accès').getDataRange().getValues(), tz = Session.getScriptTimeZone();
  var p = pushSheet_().getDataRange().getValues().slice(1), notif = {};
  p.forEach(function (r) { if (r[1] && (r[7] === true || r[7] === 'TRUE')) { var n = notif[r[1]] = notif[r[1]] || { nb: 0, app: [] }; n.nb++; if (r[6]) n.app.push(String(r[6])); } });
  return v.slice(1).filter(function (r) { return r[0] && r[2]; }).map(function (r) {
    var d = r[8] instanceof Date ? r[8] : null;
    return { nom: String(r[0]), role: String(r[1]), email: !!String(r[3] || '').trim(), mdp: !!String(r[7] || '').trim(),
      der: d ? Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm') : '', nb: Number(r[9]) || 0, notif: notif[r[0]] ? notif[r[0]].nb : 0,
      app: notif[r[0]] ? notif[r[0]].app.map(appareil_).filter(function (x, i, a) { return a.indexOf(x) === i; }) : [] };
  });
}
function appareil_(ua) { ua = String(ua || ''); return /iPhone/.test(ua) ? 'iPhone' : /iPad/.test(ua) ? 'iPad' : /Android/.test(ua) ? 'Android' : /Macintosh/.test(ua) ? 'Mac' : /Windows/.test(ua) ? 'PC' : 'autre'; }
/* Suivi d'utilisation (Guy) : connexions, appareils notifiés, wellness et RPE des 7 derniers jours, dispos de la semaine prochaine, notifications reçues. */
function getUtilisation(key) {
  var L = getConnexions(key), td = fmt_(new Date(), 'yyyy-MM-dd'), w = wellness_(8), r = rpe_(8), lun = fmt_(lundi_(1), 'yyyy-MM-dd'), dsp = dispos_(lun).sem;
  var TT = tireurs_(), PL = profilsLecons_();
  var ns = notifSheet_(), last = ns.getLastRow(), notifs = {};
  if (last > 1) { var deb = Math.max(2, last - 1500); ns.getRange(deb, 1, last - deb + 1, 7).getValues().forEach(function (x) { if (!(x[1] instanceof Date) || Date.now() - x[1].getTime() > 7 * 864e5) return; var k = norm_(x[2]), o = notifs[k] = notifs[k] || { env: 0, lues: 0 }; o.env++; if (x[6]) o.lues++; }); }
  return L.map(function (x) {
    var kn = norm_(x.nom), t = TT.filter(function (y) { return y.cle === kn || norm_(y.prenom + ' ' + y.nom) === kn; })[0] || null, k = t ? t.cle : kn, wl = w.filter(function (y) { return norm_(y.tireur) === k; }), rl = r.filter(function (y) { return norm_(y.tireur) === k; });
    var p = monProfil_({ nom: x.nom, role: x.role }, PL);
    x.pole = !!(t && t.pole); x.statut = t ? t.statut : '';
    x.well = wl.map(function (y) { return y.d; }); x.wellAuj = wl.some(function (y) { return y.d === td; });
    x.rpe = rl.length; x.rpeDer = rl.length ? rl.map(function (y) { return y.d; }).sort().pop() : '';
    x.dispos = !!(p && dsp[p.nom]); x.notifs = notifs[norm_(x.nom)] || { env: 0, lues: 0 };
    return x;
  });
}
/* Groupe « PE INSEP » (partenaires d'entraînement INSEP) : colonne AA de l'onglet Tireurs. */
function setPeInsep(key, nom, oui) {
  staff_(key);
  var sh = sheet_(IDS.BASE, 'Tireurs');
  if (String(sh.getRange(1, 27).getValue()).trim() !== 'PE INSEP') sh.getRange(1, 27).setValue('PE INSEP').setFontWeight('bold');
  var k = norm_(nom), v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) if (norm_(v[i][0] + ' ' + v[i][1]) === k) { sh.getRange(i + 2, 27).setValue(oui ? 'Oui' : ''); cacheClear_(); return 'ok'; }
  throw new Error('Tireur introuvable : ' + nom);
}
function pushSheet_() {
  var ss = ss_(IDS.BASE), sh = ss.getSheetByName('Push');
  if (!sh) { sh = ss.insertSheet('Push'); sh.getRange(1, 1, 1, 8).setValues([['Date', 'Nom', 'Rôle', 'Endpoint', 'Clé p256dh', 'Clé auth', 'Appareil', 'Actif']]); sh.setFrozenRows(1); }
  return sh;
}
function notifSheet_() {
  var ss = ss_(IDS.BASE), sh = ss.getSheetByName('Notifications');
  if (!sh) { sh = ss.insertSheet('Notifications'); sh.getRange(1, 1, 1, 7).setValues([['ID', 'Date', 'Destinataire', 'Titre', 'Message', 'Lien', 'Lu']]); sh.setFrozenRows(1); }
  return sh;
}
function pushAbonner(key, sub, appareil) {
  var u = user_(key); if (!u) throw new Error('Lien invalide');
  if (!sub || !sub.endpoint) throw new Error('Abonnement invalide');
  evt_('cloche', u);
  var sh = pushSheet_(), l = trouveLigne_(sh, 4, sub.endpoint), row = [new Date(), u.nom, u.role, sub.endpoint, (sub.keys || {}).p256dh || '', (sub.keys || {}).auth || '', String(appareil || '').slice(0, 120), true];
  if (l) sh.getRange(l, 1, 1, 8).setValues([row]); else sh.appendRow(row);
  return 'ok';
}
function pushDesabonner(key, endpoint) {
  if (!user_(key)) throw new Error('Lien invalide');
  var sh = pushSheet_(), l = trouveLigne_(sh, 4, endpoint); if (l) sh.getRange(l, 8).setValue(false);
  return 'ok';
}
function abonnes_() {
  var sh = pushSheet_(), last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 8).getValues().map(function (x, i) { return { ligne: i + 2, nom: x[1], role: x[2], endpoint: x[3], actif: x[7] === true || x[7] === 'TRUE' }; }).filter(function (x) { return x.actif && x.endpoint; });
}
/* Envoie une notification aux personnes listées (noms tels que dans l'onglet Accès). */
function pushEnvoyer_(noms, titre, corps, lien) {
  if (!noms.length) return 0;
  var sh = notifSheet_(), now = new Date(), rows = noms.map(function (n, i) { return ['N' + now.getTime().toString(36) + i, now, n, titre, corps, lien || '', '']; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  var k = {}; noms.forEach(function (n) { k[norm_(n)] = 1; });
  var subs = abonnes_().filter(function (s) { return k[norm_(s.nom)]; }), reqs = [];
  subs.forEach(function (s) {
    var aud = s.endpoint.match(/^https:\/\/[^\/]+/)[0];
    reqs.push({ url: s.endpoint, method: 'post', muteHttpExceptions: true, payload: '', headers: { Authorization: 'vapid t=' + vapidJwt_(aud) + ', k=' + vapidCles_().pub, TTL: '86400', Urgency: 'high' } });
  });
  if (!reqs.length) return 0;
  var rep = UrlFetchApp.fetchAll(reqs), shP = pushSheet_(), ok = 0;
  rep.forEach(function (r, i) { var c = r.getResponseCode(); if (c === 404 || c === 410) shP.getRange(subs[i].ligne, 8).setValue(false); else if (c < 300) ok++; });
  return ok;
}
/* Appelé par le téléphone à la réception d'une notification : renvoie les messages non lus (48 h) et les marque lus. */
function mesNotifs(key) {
  var u = user_(key); if (!u) return [];
  var sh = notifSheet_(), last = sh.getLastRow(); if (last < 2) return [];
  var deb = Math.max(2, last - 300), v = sh.getRange(deb, 1, last - deb + 1, 7).getValues(), lim = Date.now() - 48 * 3600e3, out = [], k = norm_(u.nom);
  v.forEach(function (x, i) {
    if (norm_(x[2]) !== k || !(x[1] instanceof Date) || x[1].getTime() < lim) return;
    out.push({ id: x[0], titre: x[3], corps: x[4], lien: x[5], t: x[1].getTime() }); if (!x[6]) sh.getRange(deb + i, 7).setValue(new Date());
  });
  /* chaque appareil garde la liste de ce qu'il a déjà affiché : on renvoie les messages des 48 h, du plus récent au plus ancien */
  out.sort(function (a, b) { return b.t - a.t; });
  return out.slice(0, 15);
}
/* ---- Notifications non lues (récap sur l'écran d'accueil) : colonne H « Ouvert » ---- */
function prenomDe_(nom) {
  var m = String(nom || '').trim().split(/\s+/);
  if (m.length < 2) return m[0] || '';
  /* « NOM Prénom » (tireurs) ou « Prénom Nom » (cadres) */
  if (m[0] === m[0].toUpperCase() && m[0] !== m[0].toLowerCase()) { var p = m.filter(function (x) { return x !== x.toUpperCase(); }); return p.join(' ') || m[m.length - 1]; }
  return m[0];
}
function notifOuvertCol_(sh) { if (String(sh.getRange(1, 8).getValue()) !== 'Ouvert') sh.getRange(1, 8).setValue('Ouvert'); return 8; }
function mesNonLues(key) {
  var u = user_(key); if (!u) return [];
  var sh = notifSheet_(), last = sh.getLastRow(); if (last < 2) return [];
  notifOuvertCol_(sh);
  var deb = Math.max(2, last - 400), v = sh.getRange(deb, 1, last - deb + 1, 8).getValues(), lim = Date.now() - 7 * 864e5, out = [], k = norm_(u.nom);
  v.forEach(function (x) {
    if (norm_(x[2]) !== k || !(x[1] instanceof Date) || x[1].getTime() < lim || x[7]) return;
    out.push({ id: String(x[0]), titre: x[3], corps: x[4], lien: x[5], t: x[1].getTime() });
  });
  out.sort(function (a, b) { return b.t - a.t; });
  return out.slice(0, 20);
}
function marquerLues(key, ids) {
  var u = user_(key); if (!u) throw new Error('Lien invalide');
  ids = (ids || []).map(String); if (!ids.length) return 0;
  var sh = notifSheet_(), last = sh.getLastRow(); if (last < 2) return 0;
  notifOuvertCol_(sh);
  var deb = Math.max(2, last - 400), n = last - deb + 1, v = sh.getRange(deb, 1, n, 3).getValues(), h = sh.getRange(deb, 8, n, 1).getValues(), k = norm_(u.nom), nb = 0, now = new Date();
  v.forEach(function (x, i) { if (norm_(x[2]) === k && !h[i][0] && (ids[0] === '*' || ids.indexOf(String(x[0])) >= 0)) { h[i][0] = now; nb++; } });
  if (nb) sh.getRange(deb, 8, n, 1).setValues(h);
  return nb;
}
function pushTest(key) {
  var u = user_(key); if (!u) throw new Error('Lien invalide');
  var n = pushEnvoyer_([u.nom], 'Notifications activées ✓', 'Tu recevras ici les rappels du pôle.', '');
  return n ? 'Notification envoyée à ' + n + ' appareil' + (n > 1 ? 's' : '') + '.' : 'Aucun appareil abonné pour l\'instant.';
}

/* ---- RPE (côté tireur) ---- */
function rpeSheet_() { return sheet_(IDS.BASE, 'RPE séances'); }
function getMesSeances(key) {
  var u = user_(key); if (!u) throw new Error('Lien invalide');
  var td = fmt_(new Date(), 'yyyy-MM-dd'), s = semaine_(0), j = s ? s.jours.filter(function (x) { return x.date === td; })[0] : null;
  var cr = j ? j.creneaux.filter(function (c) { return !/^LE[CÇ]ONS?$/i.test(String(c.quoi).trim()) && !/r[ée]cup|repos/i.test(c.quoi); }).map(function (c) { return { h: c.h, quoi: c.quoi, lieu: c.lieu }; }) : [];
  var sh = rpeSheet_(), last = sh.getLastRow(), deja = [];
  if (last > 1) sh.getRange(Math.max(2, last - 200), 1, Math.min(last - 1, 201), 7).getValues().forEach(function (x) { if (x[0] instanceof Date && fmt_(x[0], 'yyyy-MM-dd') === td && norm_(x[1]) === norm_(u.nom)) deja.push({ seance: x[2], duree: x[3], rpe: x[4], comm: x[6] }); });
  return { date: td, creneaux: cr, deja: deja };
}
function saveRPE(key, f) {
  var u = user_(key); if (!u || u.role === 'Invité') throw new Error('Lien invalide');
  var nom = u.role === 'Tireur' ? u.nom : f.tireur, rpe = Number(f.rpe), du = Number(f.duree) || '';
  if (!(rpe >= 0 && rpe <= 10)) throw new Error('Le RPE doit être entre 0 et 10');
  rpeSheet_().appendRow([new Date((f.date || fmt_(new Date(), 'yyyy-MM-dd')) + 'T12:00:00'), nom, f.seance || '', du, rpe, du ? du * rpe : '', f.comm || '', u.nom]);
  if (u.role === 'Tireur') try {
    pushEnvoyer_(SUIVI_ADMIN, 'RPE · ' + prenomDe_(nom) + ' ' + rpe + '/10', String(f.seance || 'Séance') + (du ? ' · ' + du + ' min · charge ' + du * rpe : '') + (f.comm ? ' — ' + f.comm : ''), '#jour');
  } catch (e) { }
  return 'ok';
}

/* ---- rappels automatiques ---- */
/* Qui reçoit la notification d'appel : Guy pour les séances collectives, Mathieu pour la prépa physique (noms de l'onglet Accès). */
var APPEL_COLLECTIF = ['Guy'], APPEL_PP = ['Mathieu Luis'];
function nomsRole_(roles) { return acces_().filter(function (a) { return roles.indexOf(a.role) >= 0; }).map(function (a) { return a.nom; }); }
/* Mercredi 20 h : dispos et souhaits de leçons de la semaine suivante (seulement à ceux qui n'ont pas rempli). */
function notifDispos() {
  var lun = fmt_(lundi_(1), 'yyyy-MM-dd'), d = dispos_(lun);
  var noms = acces_().filter(function (a) { if (['Tireur', 'Coach', 'Staff'].indexOf(a.role) < 0) return false; var p = monProfil_(a); return !(p && d.sem[p.nom]); }).map(function (a) { return a.nom; });
  return pushEnvoyer_(noms, 'Dispos et leçons de la semaine prochaine', 'Indique tes créneaux et tes souhaits de leçons pour la semaine du ' + lun.split('-').reverse().join('/') + '.', '#dispos');
}
/* Toutes les 15 min : appel au début de chaque séance (staff), RPE à la fin (tireurs présents). */
/* ---- Résumé d'activité pour Guy : regroupé, envoyé à midi et à 21 h ---- */
var SUIVI_ADMIN = ['Guy'];
var EVT_LIB = { dispos: '📅 dispos', rpe: '💪 RPE', wellness: '🌡 wellness', lecon: '📝 leçons notées', muscu: '🏋️ muscu', cloche: '🔔 cloche activée', premiere: '👋 1re connexion' };
function prenomDe_(n) { n = String(n || '').trim(); var w = n.split(/\s+/), p = w.filter(function (x) { return /[a-zà-ÿ]/.test(x); }); return p.length ? p[0] : (w[w.length - 1] || n); }
function evt_(type, u, qui) {
  try {
    if (!u || SUIVI_ADMIN.indexOf(u.nom) >= 0) return;
    var lock = LockService.getScriptLock(); if (!lock.tryLock(3000)) return;
    try {
      var p = PropertiesService.getScriptProperties(), l = JSON.parse(p.getProperty('EVTS') || '[]');
      l.push([type, String(qui || u.nom), Date.now()]); if (l.length > 300) l = l.slice(-300);
      p.setProperty('EVTS', JSON.stringify(l));
    } finally { lock.releaseLock(); }
  } catch (e) { }
}
function resumeActivite_() {
  /* Deux fois par jour seulement : à midi et à 21 h. */
  var now = new Date(), h = Number(fmt_(now, 'H')), slot = fmt_(now, 'yyyy-MM-dd') + '|' + h;
  if (h !== 12 && h !== 21) return;
  var p = PropertiesService.getScriptProperties(); if (p.getProperty('EVTS_CRENEAU') === slot) return;
  var l = JSON.parse(p.getProperty('EVTS') || '[]'); if (!l.length) return;
  p.setProperty('EVTS_CRENEAU', slot);
  var g = {}, ordre = ['dispos', 'rpe', 'lecon', 'wellness', 'muscu', 'premiere', 'cloche'];
  l.forEach(function (e) { var x = g[e[0]] = g[e[0]] || []; var n = /→/.test(e[1]) ? e[1] : prenomDe_(e[1]); if (x.indexOf(n) < 0) x.push(n); });
  var parts = ordre.concat(Object.keys(g).filter(function (k) { return ordre.indexOf(k) < 0; })).filter(function (k) { return g[k]; }).map(function (k) {
    var n = g[k]; return (EVT_LIB[k] || k) + ' : ' + (n.length <= 4 ? n.join(', ') : n.slice(0, 3).join(', ') + ' +' + (n.length - 3));
  });
  pushEnvoyer_(SUIVI_ADMIN, 'Activité du pôle · ' + l.length + ' nouveauté' + (l.length > 1 ? 's' : ''), parts.join(' · '), '#jour');
  p.setProperty('EVTS', '[]'); p.setProperty('EVTS_DERNIER', String(now.getTime()));
}
function notifSeances() {
  try { resumeActivite_(); } catch (e) { }
  try { syncMaitreAuto_(); } catch (e) { Logger.log('sync : ' + e); }
  var now = new Date(), td = fmt_(now, 'yyyy-MM-dd'), mn = Number(fmt_(now, 'H')) * 60 + Number(fmt_(now, 'm'));
  if (mn >= 9 * 60 && mn < 10 * 60) { try { envoyerRappelsWellness(); } catch (e) { Logger.log('wellness : ' + e); } }
  if (mn < 7 * 60 || mn > 22 * 60) return;
  var s = semaine_(0), j = s ? s.jours.filter(function (x) { return x.date === td; })[0] : null; if (!j) return;
  var pr = PropertiesService.getScriptProperties(), envoye = JSON.parse(pr.getProperty('NOTIF_SEANCES') || '{}');
  Object.keys(envoye).forEach(function (k) { if (k.slice(0, 10) < td) delete envoye[k]; });
  j.creneaux.forEach(function (c) {
    if (/^LE[CÇ]ONS?$/i.test(String(c.quoi).trim()) || /r[ée]cup|repos/i.test(c.quoi)) return;
    var hh = String(c.h).match(/(\d{1,2})\s*h\s*(\d{0,2})(?:\D+(\d{1,2})\s*h\s*(\d{0,2}))?/i); if (!hh) return;
    var deb = Number(hh[1]) * 60 + (Number(hh[2]) || 0), fin = hh[3] ? Number(hh[3]) * 60 + (Number(hh[4]) || 0) : deb + 90;
    var kA = td + '|A|' + deb, kR = td + '|R|' + deb, lib = String(c.h).replace(/\s+/g, '') + ' · ' + String(c.quoi).replace(/\s+/g, ' ').trim();
    if (!envoye[kA] && mn >= deb - 10 && mn <= deb + 20) { pushEnvoyer_(/\bPP\b|PR[ÉE]PA|PHYSIQUE|MUSCU|RENFO/i.test(c.quoi) ? APPEL_PP : APPEL_COLLECTIF, 'Appel · ' + lib, 'La séance commence : touche pour faire l\'appel.', '#jour'); envoye[kA] = 1; }
    if (!envoye[kR] && mn >= fin && mn <= fin + 45) {
      var pres = [], sh = sheet_(IDS.BASE, 'Appel'), se = seancesSemaine_(s.lundi).filter(function (x) { return x.date === td && Math.abs(hmin_(x.h) - deb) <= 30; })[0];
      if (se) pres = se.appel.filter(function (a) { return /^Présent|Retard|Écourté|Adapté/.test(a.statut); }).map(function (a) { return a.tireur; });
      var tir = acces_().filter(function (a) { return a.role === 'Tireur'; }), noms = [];
      tir.forEach(function (a) { var t = trouveTireur_(a.nom); if (!pres.length ? (t && t.pole) : pres.some(function (p) { return norm_(p) === (t ? t.cle : norm_(a.nom)); })) noms.push(a.nom); });
      pushEnvoyer_(noms, 'RPE · ' + lib, 'Note l\'intensité de ta séance (30 secondes).', '#rpe');
      envoye[kR] = 1;
    }
  });
  pr.setProperty('NOTIF_SEANCES', JSON.stringify(envoye));
}
/** À lancer une fois depuis l'éditeur : installe les rappels push (dispos le mercredi 20 h, appel et RPE toutes les 15 min). */
function installerNotifications() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (['notifDispos', 'notifSeances'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('notifDispos').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(20).nearMinute(0).inTimezone(TZ).create();
  ScriptApp.newTrigger('notifSeances').timeBased().everyMinutes(15).create();
  pushSheet_(); notifSheet_();
  return ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
}

/* ---- Envoi du planning des leçons validées (notification à chaque tireur et à chaque maître d'armes) ---- */
function profilDe_(a) {
  var p = monProfil_(a); if (p) return p;
  var n = norm_(a.nom);
  return profilsLecons_().filter(function (q) { return q.role === 'Coach' && (n === norm_(q.nom) || n.split(' ').indexOf(norm_(q.nom)) >= 0 || (q.prenom && n.split(' ').indexOf(norm_(q.prenom)) >= 0)); })[0] || null;
}
function envoyerPlanningLecons(key, lundi) {
  staff_(key);
  var L = leconsPlanifiees_(lundi).filter(function (l) { return /valid/i.test(l.statut); });
  if (!L.length) throw new Error('Aucune leçon validée pour cette semaine.');
  var J = ['dim.', 'lun.', 'mar.', 'mer.', 'jeu.', 'ven.', 'sam.'], sem = lundi.split('-').reverse().slice(0, 2).join('/');
  function quand(l) { var d = new Date(l.date + 'T12:00:00'); return J[d.getDay()] + ' ' + l.h; }
  L.sort(function (a, b) { return (a.date + hmin_(a.h)) < (b.date + hmin_(b.h)) ? -1 : 1; });
  var personnes = 0, appareils = 0, sans = [];
  acces_().forEach(function (a) {
    if (['Tireur', 'Coach', 'Staff'].indexOf(a.role) < 0) return;
    var p = profilDe_(a); if (!p) return;
    var k = norm_(p.nom), txt;
    if (p.role === 'Coach') {
      var m = L.filter(function (l) { return norm_(l.coach) === k; }); if (!m.length) return;
      txt = m.map(function (l) { return quand(l) + ' ' + l.tireur + (l.vs ? ' + ' + l.vs : ''); }).join(' · ');
    } else {
      var t = L.filter(function (l) { return norm_(l.tireur) === k || norm_(l.vs) === k; }); if (!t.length) return;
      txt = t.map(function (l) { return quand(l) + ' avec ' + l.coach + (l.vs ? ' (doublette)' : ''); }).join(' · ');
    }
    var n = pushEnvoyer_([a.nom], 'Leçons · semaine du ' + sem, txt, '#semaine');
    personnes++; appareils += n; if (!n) sans.push(a.nom);
  });
  return 'Planning envoyé à ' + personnes + ' personne' + (personnes > 1 ? 's' : '') + ' (' + appareils + ' téléphone' + (appareils > 1 ? 's' : '') + ' avec l\'application).' + (sans.length ? ' Pas encore l\'application : ' + sans.join(', ') + '.' : '');
}
/* Rappel ponctuel des dispos (une seule fois, à la date donnée). */
function disposJeudi9h() {
  ScriptApp.newTrigger('notifDispos').timeBased().at(new Date('2026-11-19T09:00:00+01:00')).create();
  return ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
}

/* ---- « Recevoir mon lien par e-mail » : envoie le lien personnel à l'adresse enregistrée dans Accès (colonne D). ---- */
function envoyerMonLien(email) {
  email = String(email || '').trim().toLowerCase();
  var rep = 'Si cette adresse est enregistrée au pôle, un e-mail avec ton lien personnel vient de partir (pense à regarder dans les indésirables).';
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Adresse e-mail invalide');
  var c = CacheService.getScriptCache(), ck = 'lien_' + Utilities.base64EncodeWebSafe(email).slice(0, 200);
  if (c.get(ck)) return rep;
  c.put(ck, '1', 600);
  var a = acces_().filter(function (x) { return x.email === email && x.cle; })[0];
  if (!a) { Utilities.sleep(800); return rep; }
  var app = 'https://pole-france-releve-fh.github.io/', lien = app + '?k=' + encodeURIComponent(a.cle), web = ScriptApp.getService().getUrl() + '?k=' + encodeURIComponent(a.cle);
  MailApp.sendEmail({ to: email, name: 'Pôle France Relève fleuret H', subject: 'Ton accès au portail France Fleuret Hommes',
    htmlBody: 'Bonjour ' + prenomDe_(a.nom) + ',<br><br>Voici ton lien personnel : <a href="' + lien + '">ouvrir le portail</a>.<br><br>' +
      'Une fois ouvert, touche la clé 🔑 en haut à droite pour choisir ton mot de passe : tu pourras ensuite te connecter avec ton e-mail et ce mot de passe, notamment dans l\'application installée sur ton téléphone (Safari : Partager → « Sur l\'écran d\'accueil » ; Chrome : menu ⋮ → « Ajouter à l\'écran d\'accueil »).<br><br>' +
      'Ce lien est personnel : ne le transfère pas.<br><br>Pôle France Relève fleuret H<br><span style="color:#888;font-size:12px">Lien de secours (version web) : <a href="' + web + '">' + web.replace(/\?.*/, '') + '</a></span>' });
  return rep;
}

/* =========================================================
 * Musculation : mouvements (onglet « Mouvements muscu », modifiable par Mathieu) et charges déclarées
 * (onglet « Charges muscu » : Date | Tireur | Mouvement | Charge (kg) | Répétitions | 1RM estimé | Commentaire | Saisi par)
 * 1RM estimé = charge × (1 + répétitions / 30) (formule d'Epley), = charge si 1 répétition.
 * ======================================================= */
var MUSCU_DEFAUT = [
  ['Haltérophilie', 'Épaulé (Clean)'], ['Haltérophilie', 'Épaulé suspendu (Hang Clean)'], ['Haltérophilie', 'Épaulé fente (Split Clean)'], ['Haltérophilie', 'Arraché (Snatch)'], ['Haltérophilie', 'Arraché suspendu (Hang Snatch)'], ['Haltérophilie', 'Jeté fente (Split Jerk)'], ['Haltérophilie', 'Push Press'],
  ['Bas du corps', 'Squat'], ['Bas du corps', 'Front Squat'], ['Bas du corps', 'Box Squat'], ['Bas du corps', 'Soulevé de terre roumain (RDL)'], ['Bas du corps', 'Fentes arrière / bulgares (charge totale)'], ['Bas du corps', 'Hip Thrust'],
  ['Haut du corps · poussée', 'Développé couché (Bench Press)'], ['Haut du corps · poussée', 'Développé militaire (Strict Press)'], ['Haut du corps · poussée', 'Dips lestés (charge ajoutée)'], ['Haut du corps · poussée', 'Landmine Press (1 bras)'],
  ['Haut du corps · tirage', 'Tractions lestées (charge ajoutée)'], ['Haut du corps · tirage', 'Tractions au poids du corps (répétitions max)'], ['Haut du corps · tirage', 'Rowing barre (Bent Over / Pendlay Row)'], ['Haut du corps · tirage', 'Seal Row'], ['Haut du corps · tirage', 'Rowing haltère (1 bras)']
];
function muscuMouvements_() {
  var ss = ss_(IDS.BASE), sh = ss.getSheetByName('Mouvements muscu');
  if (!sh) {
    sh = ss.insertSheet('Mouvements muscu');
    sh.getRange(1, 1, 1, 3).setValues([['Famille', 'Mouvement', 'Actif (OUI/NON)']]);
    sh.getRange(2, 1, MUSCU_DEFAUT.length, 3).setValues(MUSCU_DEFAUT.map(function (m) { return [m[0], m[1], 'OUI']; }));
    sh.setFrozenRows(1);
  }
  return sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).getDisplayValues().filter(function (x) { return x[1] && String(x[2]).toUpperCase() !== 'NON'; }).map(function (x) { return { famille: x[0], nom: x[1] }; });
}
function chargesSheet_() {
  var ss = ss_(IDS.BASE), sh = ss.getSheetByName('Charges muscu');
  if (!sh) { sh = ss.insertSheet('Charges muscu'); sh.getRange(1, 1, 1, 8).setValues([['Date', 'Tireur', 'Mouvement', 'Charge (kg)', 'Répétitions', '1RM estimé (kg)', 'Commentaire', 'Saisi par']]); sh.setFrozenRows(1); }
  return sh;
}
function e1rm_(c, r) { c = Number(c) || 0; r = Number(r) || 0; if (!c) return 0; return Math.round((r <= 1 ? c : c * (1 + r / 30)) * 10) / 10; }
function charges_() {
  var sh = chargesSheet_(), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 8).getValues().filter(function (x) { return x[1] && x[2]; }).map(function (x, i) {
    return { ligne: i + 2, date: x[0] instanceof Date ? fmt_(x[0], 'yyyy-MM-dd') : String(x[0]), tireur: String(x[1]), mvt: String(x[2]), charge: Number(x[3]) || 0, reps: Number(x[4]) || 0, e1rm: Number(x[5]) || e1rm_(x[3], x[4]), comm: x[6], par: x[7] };
  });
}
function muscuNom_(u, nom) {
  if (!u || u.role === 'Invité' || u.role === 'Sélectionneur') throw new Error('Accès non autorisé');
  if (u.role === 'Staff') return nom;
  if (u.role === 'Tireur') return u.nom;
  throw new Error('Accès non autorisé');
}
function getMuscu(key, nom) {
  var u = user_(key), qui = muscuNom_(u, nom), t = trouveTireur_(qui), k = t ? t.cle : norm_(qui);
  var mien = charges_().filter(function (c) { var tt = trouveTireur_(c.tireur); return (tt ? tt.cle : norm_(c.tireur)) === k; });
  mien.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  return { tireur: qui, mouvements: muscuMouvements_(), entrees: mien.slice(0, 300) };
}
function saveCharge(key, f) {
  var u = user_(key), qui = muscuNom_(u, f.tireur);
  var c = Number(String(f.charge).replace(',', '.')), r = Math.round(Number(f.reps));
  if (!f.mvt) throw new Error('Choisis le mouvement');
  if (!(c >= 0 && c <= 400)) throw new Error('Charge invalide (en kg)');
  evt_('muscu', u, qui);
  if (!(r >= 1 && r <= 50)) throw new Error('Nombre de répétitions invalide');
  var t = trouveTireur_(qui), nomT = t ? t.nom + ' ' + t.prenom : qui;
  chargesSheet_().appendRow([new Date((f.date || fmt_(new Date(), 'yyyy-MM-dd')) + 'T12:00:00'), nomT, f.mvt, c, r, e1rm_(c, r), f.comm || '', u.nom]);
  return 'ok';
}
function supprimerCharge(key, ligne, mvt, date) {
  var u = user_(key), sh = chargesSheet_(); ligne = Number(ligne); if (!(ligne >= 2 && ligne <= sh.getLastRow())) throw new Error('Ligne introuvable');
  var v = sh.getRange(ligne, 1, 1, 3).getValues()[0];
  var d = v[0] instanceof Date ? fmt_(v[0], 'yyyy-MM-dd') : String(v[0]);
  if (String(v[2]) !== mvt || d !== date) throw new Error('Ligne modifiée entre-temps, recharge la page');
  if (!u) throw new Error('Lien invalide');
  if (u.role !== 'Staff' && !(u.role === 'Tireur' && norm_(v[1]).indexOf(norm_(u.nom).split(' ')[0]) >= 0)) throw new Error('Accès non autorisé');
  sh.deleteRow(ligne);
  return 'ok';
}
/* Vue d'ensemble pour le staff : meilleur 1RM estimé par tireur et par mouvement. */
function getMuscuEquipe(key) {
  staff_(key);
  var best = {}, mv = {};
  charges_().forEach(function (c) {
    var t = trouveTireur_(c.tireur), n = t ? t.nom + ' ' + t.prenom : c.tireur, b = best[n] = best[n] || {};
    mv[c.mvt] = 1;
    if (!b[c.mvt] || c.e1rm > b[c.mvt].e1rm) b[c.mvt] = { e1rm: c.e1rm, charge: c.charge, reps: c.reps, date: c.date };
  });
  var ordre = muscuMouvements_().map(function (m) { return m.nom; }).filter(function (m) { return mv[m]; });
  Object.keys(mv).forEach(function (m) { if (ordre.indexOf(m) < 0) ordre.push(m); });
  return { mouvements: ordre, best: best, tireurs: tireurs_().filter(function (t) { return t.pole || /^Partenaire/i.test(t.statut); }).map(function (t) { return { nom: t.nom + ' ' + t.prenom, pole: t.pole }; }) };
}
function lancerEtape1() { return lancerPlanningApp('2026-11-23', 2, '2026-11-30'); }
function lancerEtape2() { return importerCompets(); }
