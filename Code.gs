// ╔══════════════════════════════════════════════════════════════════════╗
// ║          COUPE DU MONDE 2026 — Google Apps Script                   ║
// ║  Auteur : matvanghe2-boop                                           ║
// ║  Repo   : https://github.com/matvanghe2-boop/CDM2026                ║
// ║                                                                      ║
// ║  Ce script lit l'onglet "Matchs" du Google Sheet CDM_2026_Live,     ║
// ║  génère scores.json et groups.json, puis les pousse directement      ║
// ║  sur GitHub via l'API. Vercel détecte le commit et redéploie.       ║
// ╚══════════════════════════════════════════════════════════════════════╝

// ──────────────────────────────────────────────────────────────────────
//  ① CONFIGURATION — à remplir une seule fois
// ──────────────────────────────────────────────────────────────────────
var CONFIG = {
  // Token GitHub : Settings → Developer settings → Personal access tokens → Fine-grained
  // Permissions nécessaires : Contents (Read & Write) sur le repo CDM2026
  GITHUB_TOKEN : 'COLLE_TON_TOKEN_ICI',

  // Coordonnées du dépôt
  GITHUB_OWNER : 'matvanghe2-boop',
  GITHUB_REPO  : 'CDM2026',
  GITHUB_BRANCH: 'main',         // ou 'master' selon ton repo

  // Noms des fichiers à créer/écraser à la racine du repo
  FILE_SCORES : 'scores.json',
  FILE_GROUPS : 'groups.json',

  // Nom de l'onglet dans le Google Sheet
  SHEET_NAME  : 'Matchs',
};

// ──────────────────────────────────────────────────────────────────────
//  ② POINT D'ENTRÉE PRINCIPAL
//  C'est cette fonction qu'on branche sur le déclencheur temporel.
// ──────────────────────────────────────────────────────────────────────
function exporterVersGitHub() {
  try {
    Logger.log('▶ Démarrage export CDM 2026…');

    var matches = lireMatchs();
    Logger.log('  Matchs lus : ' + matches.length);

    var groups  = calculerClassements(matches);

    var scoresJson = JSON.stringify(matches, null, 2);
    var groupsJson = JSON.stringify(groups,  null, 2);

    var sha_scores = getSHA(CONFIG.FILE_SCORES);
    var sha_groups = getSHA(CONFIG.FILE_GROUPS);

    pusherFichier(CONFIG.FILE_SCORES, scoresJson, sha_scores,
                  '🤖 scores.json mis à jour — ' + horodatage());
    pusherFichier(CONFIG.FILE_GROUPS, groupsJson, sha_groups,
                  '🤖 groups.json mis à jour — ' + horodatage());

    Logger.log('✅ Export terminé avec succès.');
  } catch (e) {
    Logger.log('❌ Erreur : ' + e.message);
    // En cas d'erreur répétée, tu recevras un e-mail si tu as activé
    // les notifications d'erreur dans les paramètres Apps Script.
    throw e;
  }
}

// ──────────────────────────────────────────────────────────────────────
//  ③ LECTURE DE L'ONGLET "Matchs"
//  Lit les 10 colonnes dans l'ordre : id | group | home | away |
//  hs | as | date | finished | phase | type
// ──────────────────────────────────────────────────────────────────────
function lireMatchs() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Onglet "' + CONFIG.SHEET_NAME + '" introuvable dans le tableur.');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];  // Aucune donnée (ligne 1 = en-têtes)

  // Récupère toutes les lignes de données (à partir de la ligne 2)
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  var matches = [];

  data.forEach(function(row, index) {
    // Ignore les lignes complètement vides
    if (!row[0] && !row[2] && !row[3]) return;

    var id       = row[0];   // Colonne A
    var group    = String(row[1] || '').trim().toUpperCase();   // Colonne B
    var home     = String(row[2] || '').trim();  // Colonne C
    var away     = String(row[3] || '').trim();  // Colonne D
    var hs       = parseScore(row[4]);            // Colonne E (home score)
    var as_      = parseScore(row[5]);            // Colonne F (away score)
    var date     = formatDate(row[6]);            // Colonne G
    var finished = normalizeFinished(row[7]);     // Colonne H
    var phase    = String(row[8] || 'PRE').trim().toUpperCase(); // Colonne I
    var type     = String(row[9] || 'group').trim().toLowerCase(); // Colonne J

    // Validation minimale : on a besoin des deux équipes
    if (!home || !away) {
      Logger.log('  Ligne ' + (index + 2) + ' ignorée : équipes manquantes.');
      return;
    }

    matches.push({
      id      : id || (index + 1),
      group   : group,
      home    : home,
      away    : away,
      hs      : hs,
      as      : as_,
      date    : date,
      finished: finished,
      phase   : phase,
      type    : type
    });
  });

  return matches;
}

// ──────────────────────────────────────────────────────────────────────
//  ④ CALCUL DES CLASSEMENTS DE GROUPE
//  Reconstruit les standings à partir des matchs terminés.
//  Tri FIFA : pts → diff buts → buts marqués → ordre alpha
// ──────────────────────────────────────────────────────────────────────
function calculerClassements(matches) {

  // Structure de référence : les 12 groupes et leurs 4 équipes
  var SGROUPS = [
    {l:'A', t:['Mexico','South Korea','South Africa','Czechia']},
    {l:'B', t:['Canada','Switzerland','Qatar','Bosnia and Herzegovina']},
    {l:'C', t:['Brazil','Morocco','Scotland','Haiti']},
    {l:'D', t:['United States','Australia','Paraguay','Türkiye']},
    {l:'E', t:['Germany','Ecuador','Ivory Coast','Curaçao']},
    {l:'F', t:['Netherlands','Japan','Sweden','Tunisia']},
    {l:'G', t:['Belgium','Iran','Egypt','New Zealand']},
    {l:'H', t:['Spain','Uruguay','Saudi Arabia','Cape Verde']},
    {l:'I', t:['France','Senegal','Norway','Iraq']},
    {l:'J', t:['Argentina','Austria','Algeria','Jordan']},
    {l:'K', t:['Portugal','Colombia','Uzbekistan','DR Congo']},
    {l:'L', t:['England','Croatia','Panama','Ghana']},
  ];

  // Initialise toutes les équipes à zéro
  var map = {};
  SGROUPS.forEach(function(g) {
    map[g.l] = {
      group: g.l,
      teams: g.t.map(function(name) {
        return { name: name, mp:0, w:0, d:0, l:0, gf:0, ga:0, pts:0 };
      })
    };
  });

  // Parcourt les matchs terminés
  matches.forEach(function(m) {
    if (m.finished !== 'TRUE') return;
    if (m.type !== 'group') return;

    var grp = map[m.group];
    if (!grp) return;

    var hs  = Number(m.hs);
    var as_ = Number(m.as);
    if (isNaN(hs) || isNaN(as_)) return;

    // Recherche des équipes (tolérant aux variations de casse)
    var home = trouverEquipe(grp.teams, m.home);
    var away = trouverEquipe(grp.teams, m.away);

    if (!home || !away) {
      Logger.log('  Équipe non trouvée dans le groupe ' + m.group +
                 ' : "' + m.home + '" vs "' + m.away + '"');
      return;
    }

    // Compteurs communs
    home.mp++; away.mp++;
    home.gf += hs;  home.ga += as_;
    away.gf += as_; away.ga += hs;

    // Résultat
    if (hs > as_)      { home.w++; home.pts += 3; away.l++; }
    else if (hs < as_) { away.w++; away.pts += 3; home.l++; }
    else               { home.d++; home.pts++;     away.d++; away.pts++; }
  });

  // Trie chaque groupe et calcule la différence de buts
  return Object.values(map).map(function(g) {
    g.teams.sort(function(a, b) {
      var diffA = a.gf - a.ga, diffB = b.gf - b.ga;
      return (b.pts - a.pts) || (diffB - diffA) || (b.gf - a.gf) ||
             a.name.localeCompare(b.name);
    });
    return g;
  });
}

// ──────────────────────────────────────────────────────────────────────
//  ⑤ API GITHUB — push d'un fichier
// ──────────────────────────────────────────────────────────────────────

// Récupère le SHA du fichier existant (requis par l'API pour le mettre à jour).
// Retourne null si le fichier n'existe pas encore (première création).
function getSHA(filename) {
  var url = 'https://api.github.com/repos/' +
            CONFIG.GITHUB_OWNER + '/' + CONFIG.GITHUB_REPO +
            '/contents/' + filename +
            '?ref=' + CONFIG.GITHUB_BRANCH;

  var options = {
    method     : 'get',
    headers    : githubHeaders(),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code     = response.getResponseCode();

  if (code === 200) {
    var data = JSON.parse(response.getContentText());
    return data.sha;
  } else if (code === 404) {
    return null;  // Fichier inexistant → sera créé
  } else {
    throw new Error('getSHA HTTP ' + code + ' : ' + response.getContentText());
  }
}

// Crée ou écrase un fichier dans le repo via l'API GitHub Contents.
function pusherFichier(filename, content, sha, commitMessage) {
  var url = 'https://api.github.com/repos/' +
            CONFIG.GITHUB_OWNER + '/' + CONFIG.GITHUB_REPO +
            '/contents/' + filename;

  var body = {
    message: commitMessage,
    content : Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch  : CONFIG.GITHUB_BRANCH
  };

  // Si le fichier existe déjà, on doit fournir son SHA pour l'écraser
  if (sha) body.sha = sha;

  var options = {
    method     : 'put',
    headers    : githubHeaders(),
    payload    : JSON.stringify(body),
    contentType: 'application/json',
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code     = response.getResponseCode();

  if (code !== 200 && code !== 201) {
    throw new Error('pusherFichier "' + filename + '" HTTP ' + code +
                    ' : ' + response.getContentText());
  }

  Logger.log('  ✓ ' + filename + ' poussé (HTTP ' + code + ')');
}

function githubHeaders() {
  return {
    'Authorization': 'Bearer ' + CONFIG.GITHUB_TOKEN,
    'Accept'       : 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
}

// ──────────────────────────────────────────────────────────────────────
//  ⑥ GESTION DU DÉCLENCHEUR AUTOMATIQUE (toutes les 15 minutes)
// ──────────────────────────────────────────────────────────────────────

// Appelle cette fonction UNE SEULE FOIS manuellement pour installer le déclencheur.
// Elle crée un déclencheur "toutes les 15 minutes" sur exporterVersGitHub().
function installerDeclencheur() {
  // Supprime les éventuels déclencheurs existants pour éviter les doublons
  supprimerTousLesDeclencheurs();

  ScriptApp.newTrigger('exporterVersGitHub')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('✅ Déclencheur installé : exporterVersGitHub toutes les 15 minutes.');
}

// Supprime tous les déclencheurs du projet (utile pour repartir de zéro).
function supprimerTousLesDeclencheurs() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  Logger.log('🗑️ ' + triggers.length + ' déclencheur(s) supprimé(s).');
}

// ──────────────────────────────────────────────────────────────────────
//  ⑦ UTILITAIRES
// ──────────────────────────────────────────────────────────────────────

// Recherche une équipe dans la liste par nom exact ou insensible à la casse
function trouverEquipe(teams, nom) {
  var n = String(nom || '').trim().toLowerCase();
  return teams.find(function(t) {
    return t.name.toLowerCase() === n;
  }) || null;
}

// Normalise les scores : null si vide ou non numérique, nombre sinon
function parseScore(val) {
  if (val === '' || val === null || val === undefined) return null;
  var n = Number(val);
  return isNaN(n) ? null : n;
}

// Normalise le champ finished → toujours la string "TRUE" ou "FALSE"
function normalizeFinished(val) {
  if (val === true  || String(val).toUpperCase() === 'TRUE'  || val === 1) return 'TRUE';
  if (val === false || String(val).toUpperCase() === 'FALSE' || val === 0) return 'FALSE';
  return 'FALSE';
}

// Formate la date en ISO 8601 UTC (ex: "2026-06-11T19:00:00.000Z")
// Gère les objets Date de Sheets et les strings déjà formatées
function formatDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return val.toISOString();
  }
  var s = String(val).trim();
  if (!s) return null;
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : d.toISOString();
}

function horodatage() {
  return Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
}
