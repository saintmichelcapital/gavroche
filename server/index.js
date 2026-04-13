const express = require('express');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.post('/api/generate', async (req, res) => {
  try {
    const { entite, typeMission, typeDD, commentaires, complement } = req.body;
    const Anthropic = require('@anthropic-ai/sdk');
    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
    const prompt = [
      "Tu es expert en due diligence financiere pour Gavroche.",
      "Mission : " + typeMission + " (" + typeDD + ") sur " + entite.nom,
      "Siren : " + entite.siren + " | Activite : " + entite.activite,
      "Siege : " + entite.siege + " | Creation : " + entite.creation,
      "Effectifs : " + entite.effectifs,
      complement ? "Complements : " + complement : "",
      commentaires ? "Notes : " + commentaires : "",
      "Recherche des infos sur cette entreprise et genere UNIQUEMENT ce JSON sans markdown :",
      '{"contexte_operation":"","activite_resume":"","chaine_valeur":"","clients_types":"","fournisseurs_types":"","axes_analyse":["","",""],"metriques":[{"label":"CA","valeur":"","source":""},{"label":"EBITDA","valeur":"","source":""},{"label":"Effectifs","valeur":"' + (entite.effectifs||'') + '","source":"API gouvernementale"}],"marche":[{"label":"","valeur":"","source":""}],"presse":[{"date":"","titre":"","source":""}],"has_presse":false,"organigramme":{"holding":{"nom":"","pct_detention":""},"cible":{"nom":"' + (entite.nom||'') + '","perimetre":true},"filiales":[],"hors_perimetre":[]}}'
    ].filter(Boolean).join("\n");

    const response = await client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: 2000,
      tools: [{ type: "web_search_20250305", name: "web_search" }],
      messages: [{ role: "user", content: prompt }]
    });
    const text = (response.content || []).filter(b => b.type === 'text').map(b => b.text).join('').trim();
    const data = JSON.parse(text.replace(/```json|```/g, '').trim());
    res.json({ success: true, data });
  } catch (error) {
    console.error('Erreur generate:', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/chat', async (req, res) => {
  try {
    const { message, historique, contexte } = req.body;
    const Anthropic = require('@anthropic-ai/sdk');
    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
    const system = "Tu es assistant de Gavroche, cabinet de due diligence financiere parisien. " +
      "Tu aides a finaliser une proposition commerciale de " + (contexte.typeMission || 'DD') + 
      " pour " + (contexte.cible || 'la cible') + ". " +
      "Slide affichee : " + (contexte.slide || '--') + ". " +
      "Quand on te demande de modifier ou rediger un texte pour une slide, fournis DIRECTEMENT le nouveau texte exact entre guillemets, sans explication. " +
      "Tu peux reformuler, completer, raccourcir ou developper n'importe quel contenu. Ton sobre et professionnel.";
    const response = await client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: 800,
      system: system,
      messages: [...(historique || []), { role: "user", content: message }]
    });
    res.json({ success: true, reply: response.content[0]?.text || 'Erreur.' });
  } catch (error) {
    console.error('Erreur chat:', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/site-web', async (req, res) => {
  try {
    const { nomEntite } = req.body;
    const Anthropic = require('@anthropic-ai/sdk');
    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
    const response = await client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: 100,
      tools: [{ type: "web_search_20250305", name: "web_search" }],
      messages: [{ role: "user", content: "URL officielle du site de " + nomEntite + ". Reponds uniquement avec l URL ou non_trouve." }]
    });
    const text = (response.content || []).filter(b => b.type === 'text').map(b => b.text).join('').trim();
    const url = (text && text !== 'non_trouve' && text.startsWith('http')) ? text : '';
    res.json({ success: true, url });
  } catch (error) {
    res.json({ success: true, url: '' });
  }
});

app.listen(PORT, () => {
  console.log('Serveur Gavroche demarre sur http://localhost:' + PORT);
});