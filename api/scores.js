export default async function handler(request, response) {
    // On force les entêtes pour autoriser ton propre site à lire les données
    response.setHeader('Access-Control-Allow-Origin', '*');
    
    try {
        const apiRes = await fetch('https://worldcup26.ir/get/games');
        const data = await apiRes.json();
        
        // Renvoie les données en direct à ton HTML
        return response.status(200).json(data);
    } catch (error) {
        return response.status(500).json({ error: "Impossible de récupérer les scores" });
    }
}