const fs = require('fs');

// L'URL d'OpenFootball (fiable, au format JSON brut)
const URL_SOURCE = 'https://raw.githubusercontent.com/openfootball/worldcup.json/master/2026/worldcup.json';

async function scraperScores() {
  console.log("🔄 Récupération des dernières données sur Internet...");
  try {
    const response = await fetch(URL_SOURCE);
    if (!response.ok) throw new Error(`Erreur HTTP : ${response.status}`);
    
    const rawData = await response.json();
    const matches = [];
    const rounds = rawData.rounds || [];

    // Formatage des données pour ton index.html
    rounds.forEach(r => {
      if (!r.matches) return;
      r.matches.forEach(m => {
        const grp = (m.group || '').replace('Group ', '').trim();
        const finished = m.score1 !== undefined && m.score1 !== null;
        
        matches.push({
          id: m.num,
          group: grp,
          home: m.team1?.name || '',
          away: m.team2?.name || '',
          hs: finished ? Number(m.score1) : null,
          as: finished ? Number(m.score2) : null,
          date: m.date ? `${m.date}T${m.time || '19:00'}:00Z` : null,
          finished: finished ? 'TRUE' : 'FALSE',
          phase: finished ? 'FT' : 'PRE',
          type: grp ? 'group' : 'knockout'
        });
      });
    });

    // Écriture du fichier local scores.json
    fs.writeFileSync('scores.json', JSON.stringify(matches, null, 2), 'utf-8');
    console.log("✅ Fichier 'scores.json' mis à jour avec succès !");

  } catch (error) {
    console.error("❌ Le scraping a échoué :", error.message);
  }
}

scraperScores();