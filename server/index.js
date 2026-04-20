const express = require('express');
const cors = require('cors');
const path = require('path');
const jwt = require('jsonwebtoken');
const rateLimit = require('express-rate-limit');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET || 'gavroche_secret_2026';

app.use(cors({
  origin: ['https://gavroche-production.up.railway.app'],
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(express.json());
app.use(express.static(path.join(__dirname, '../public')));

// ═══════════════════════════════════════════
// RATE LIMITING
// ═══════════════════════════════════════════

// Limite generale : 100 requetes / 15 min par IP
const limiterGeneral = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 100,
  message: { success: false, error: 'Trop de requetes. Reessayez dans 15 minutes.' }
});

// Limite stricte sur les appels IA : 10 / min par IP
const limiterIA = rateLimit({
  windowMs: 60 * 1000,
  max: 10,
  message: { success: false, error: 'Limite atteinte. Reessayez dans une minute.' }
});

app.use('/api/', limiterGeneral);
app.use('/api/generate', limiterIA);
app.use('/api/chat', limiterIA);

// ═══════════════════════════════════════════
// AUTHENTIFICATION JWT
// ═══════════════════════════════════════════

// Middleware de verification du token
function requireAuth(req, res, next) {
  const auth = req.headers['authorization'];
  if (!auth || !auth.startsWith('Bearer ')) {
    return res.status(401).json({ success: false, error: 'Non autorise.' });
  }
  const token = auth.split(' ')[1];
  try {
    req.user = jwt.verify(token, JWT_SECRET);
    next();
  } catch (e) {
    return res.status(401).json({ success: false, error: 'Token invalide ou expire.' });
  }
}

// Route de login — retourne un JWT
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  const ADMIN_USER = process.env.ADMIN_USER || 'F_Thiery';
  const ADMIN_PASS = process.env.ADMIN_PASS || 'Paris75006!';
  if (username === ADMIN_USER && password === ADMIN_PASS) {
    const token = jwt.sign({ username }, JWT_SECRET, { expiresIn: '8h' });
    return res.json({ success: true, token });
  }
  return res.status(401).json({ success: false, error: 'Identifiants incorrects.' });
});

// ═══════════════════════════════════════════
// ROUTES
// ═══════════════════════════════════════════

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.post('/api/generate', requireAuth, async (req, res) => {
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

app.post('/api/chat', requireAuth, async (req, res) => {
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

app.post('/api/site-web', requireAuth, async (req, res) => {
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

// ═══════════════════════════════════════════
// EXPORT PPTX — Slide Honoraires charte Saint-Michel Capital
// ═══════════════════════════════════════════
app.post('/api/export-honoraires-pptx', requireAuth, async (req, res) => {
  try {
    const PptxGenJS = require('pptxgenjs');
    const { smcHon, cibleNom } = req.body;
    if (!smcHon) return res.status(400).json({ success: false, error: 'smcHon manquant.' });

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';              // 13.333 x 7.5 inch
    pptx.defineLayout({ name: 'A4L', width: 11.69, height: 8.27 }); // A4 paysage
    pptx.layout = 'A4L';
    pptx.title = (cibleNom || 'Propale') + ' — Honoraires';
    pptx.company = 'Saint-Michel Capital';

    const slide = pptx.addSlide();
    slide.background = { color: 'FFFFFF' };

    const W = 11.69, H = 8.27;
    const ML = 0.7, MR = 0.7, MT = 0.5, MB = 0.4;
    const COL_RIGHT_W = 2.0;   // nav droite
    const TITLE_W = W - ML - MR - COL_RIGHT_W - 0.3;

    // Titre : "Bold | suite" — pptxgenjs supporte un tableau de runs dans text
    slide.addText(
      [
        { text: smcHon.titreBold || 'Honoraires', options: { bold: true, color: '1A1A1A' } },
        { text: ' | ' + (smcHon.titreSuite || ''), options: { bold: false, color: '1A1A1A' } }
      ],
      { x: ML, y: MT, w: TITLE_W, h: 1.1, fontFace: 'Segoe UI', fontSize: 14, align: 'left', valign: 'top' }
    );

    // Nav verticale haut-droite
    const navX = W - MR - COL_RIGHT_W;
    const navItems = [
      { t: 'Contexte', active: false },
      { t: 'Périmètre des travaux', active: false },
      { t: 'Conditions', active: true }
    ];
    let navY = MT + 0.05;
    navItems.forEach(it => {
      if (it.active) {
        slide.addShape(pptx.ShapeType.rect, { x: navX, y: navY + 0.02, w: 0.03, h: 0.28, fill: { color: '1A1A1A' }, line: { color: '1A1A1A' } });
      }
      slide.addText(it.t, {
        x: navX + 0.12, y: navY, w: COL_RIGHT_W - 0.12, h: 0.3,
        fontFace: 'Segoe UI', fontSize: 8,
        color: it.active ? '1A1A1A' : '888888',
        align: 'left', valign: 'top'
      });
      navY += 0.34;
    });

    // Trait noir court sous le titre
    slide.addShape(pptx.ShapeType.rect, { x: ML, y: MT + 1.2, w: 0.8, h: 0.035, fill: { color: '1A1A1A' }, line: { color: '1A1A1A' } });

    // Corps : deux colonnes
    const BODY_Y = MT + 1.5;
    const BODY_H = H - BODY_Y - MB - 0.5;  // réserve pour footer
    const LEFT_W = (W - ML - MR) * 0.36;
    const GAP = 0.35;
    const RIGHT_X = ML + LEFT_W + GAP;
    const RIGHT_W = W - MR - RIGHT_X;

    // Colonne gauche — Prestation / Montant (HT) stacked labels
    slide.addText('Prestation', { x: ML, y: BODY_Y, w: LEFT_W, h: 0.25, fontFace: 'Segoe UI', fontSize: 10, bold: true, color: '1A1A1A' });
    slide.addText('Montant (HT)', { x: ML, y: BODY_Y + 0.25, w: LEFT_W, h: 0.25, fontFace: 'Segoe UI', fontSize: 10, bold: true, color: '1A1A1A' });

    // Prestations filtrées (montant > 0)
    let prestY = BODY_Y + 0.75;
    const prestations = (smcHon.prestations || []).filter(p => (parseInt(p.montant) || 0) > 0);
    prestations.forEach(p => {
      slide.addText(p.libelle + '  ' + '.'.repeat(40), { x: ML, y: prestY, w: LEFT_W, h: 0.28, fontFace: 'Segoe UI', fontSize: 10, color: '1A1A1A' });
      slide.addText('€ ' + (parseInt(p.montant) || 0).toLocaleString('fr-FR').replace(/\u202F/g, ' '), {
        x: ML, y: prestY + 0.32, w: 1.5, h: 0.3,
        fontFace: 'Segoe UI', fontSize: 11, color: '1A1A1A',
        fill: { color: 'D9D9D9' }, margin: 0.08
      });
      prestY += 0.8;
    });

    // Clause de réduction (optionnelle)
    if (smcHon.reductionActive && smcHon.reductionMt) {
      slide.addText("En cas de non-réalisation de l'opération envisagée, les honoraires seront réduits à € " + (parseInt(smcHon.reductionMt) || 0).toLocaleString('fr-FR').replace(/\u202F/g, ' ') + " hors taxes.",
        { x: ML, y: prestY, w: LEFT_W, h: 0.5, fontFace: 'Segoe UI', fontSize: 9, italic: true, color: '404040' });
      prestY += 0.55;
    }

    // Durée estimée + intro livrables + liste livrables
    prestY += 0.15;
    slide.addText(smcHon.dureeTxt || '', { x: ML, y: prestY, w: LEFT_W, h: 0.4, fontFace: 'Segoe UI', fontSize: 9, color: '404040' });
    prestY += 0.45;
    slide.addText(smcHon.livrIntro || '', { x: ML, y: prestY, w: LEFT_W, h: 0.4, fontFace: 'Segoe UI', fontSize: 9, color: '404040' });
    prestY += 0.4;
    (smcHon.livrables || []).forEach(l => {
      slide.addText('⊢  ' + l, { x: ML + 0.05, y: prestY, w: LEFT_W - 0.05, h: 0.35, fontFace: 'Segoe UI', fontSize: 9, color: '404040' });
      prestY += 0.38;
    });

    // Colonne droite — hypothèses à puces
    const bullets = [];
    (smcHon.hypotheses || []).forEach(h => {
      bullets.push({ text: h.txt, options: { bullet: { code: '25AA' }, fontSize: 10, color: '404040', paraSpaceAfter: 4, paraSpaceBefore: 2 } });
      (h.sub || []).filter(s => s && s.trim()).forEach(s => {
        bullets.push({ text: s, options: { bullet: { code: '22A2' }, fontSize: 10, color: '404040', indentLevel: 1, paraSpaceAfter: 3 } });
      });
    });
    slide.addText(bullets, { x: RIGHT_X, y: BODY_Y, w: RIGHT_W, h: BODY_H, fontFace: 'Segoe UI', valign: 'top' });

    // Footer : trait + S-M.C + Page X sur Y
    const footerY = H - MB;
    slide.addShape(pptx.ShapeType.line, { x: ML, y: footerY - 0.12, w: W - ML - MR, h: 0, line: { color: 'E0E0DA', width: 0.5 } });
    slide.addText('S-M.C', { x: ML, y: footerY, w: 2, h: 0.3, fontFace: 'Segoe UI', fontSize: 11, bold: true, color: '1A1A1A' });
    slide.addText('Page ' + (smcHon.pageCur || 1) + ' sur ' + (smcHon.pageTot || 1), {
      x: W - MR - 2, y: footerY, w: 2, h: 0.3,
      fontFace: 'Segoe UI', fontSize: 8, color: '888888', align: 'right'
    });

    const buf = await pptx.write({ outputType: 'nodebuffer' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="' + (cibleNom || 'Propale').replace(/[^a-zA-Z0-9_-]/g, '_') + '_Honoraires.pptx"');
    res.send(buf);
  } catch (err) {
    console.error('Erreur export PPTX:', err.message, err.stack);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.listen(PORT, () => {
  console.log('Serveur Gavroche demarre sur http://localhost:' + PORT);
});
