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

    // Titre : "Bold | suite" — aligné strictement à gauche (margin:0), interligne 1.4
    slide.addText(
      [
        { text: smcHon.titreBold || 'Honoraires', options: { bold: true, color: '1A1A1A' } },
        { text: ' | ' + (smcHon.titreSuite || ''), options: { bold: false, color: '1A1A1A' } }
      ],
      { x: ML, y: MT, w: TITLE_W, h: 1.1, fontFace: 'Segoe UI', fontSize: 13, align: 'left', valign: 'top', margin: 0, lineSpacingMultiple: 1.4 }
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

    // Trait noir court sous le titre — noir absolu, épaissi
    slide.addShape(pptx.ShapeType.rect, { x: ML, y: MT + 1.2, w: 0.8, h: 0.05, fill: { color: '000000' }, line: { color: '000000' } });

    // Corps : deux colonnes
    const BODY_Y = MT + 1.5;
    const BODY_H = H - BODY_Y - MB - 0.5;  // réserve pour footer
    const LEFT_W = (W - ML - MR) * 0.36;
    const GAP = 0.35;
    const RIGHT_X = ML + LEFT_W + GAP;
    const RIGHT_W = W - MR - RIGHT_X;

    // Colonne gauche — Prestation / Montant (HT) : vrai tableau 2 colonnes, fond blanc, taille 9, sans gras
    const MT_COL_W = 1.1;
    const LIB_COL_W = LEFT_W - MT_COL_W;
    const prestLabel = (smcHon.prestations || []).filter(p => (parseInt(p.montant) || 0) > 0).length > 1 ? 'Prestation(s)' : 'Prestation';
    slide.addText(prestLabel, { x: ML, y: BODY_Y, w: LIB_COL_W, h: 0.25, fontFace: 'Segoe UI', fontSize: 9, bold: false, color: '1A1A1A', margin: 0 });
    slide.addText('Montant (HT)', { x: ML + LIB_COL_W, y: BODY_Y, w: MT_COL_W, h: 0.25, fontFace: 'Segoe UI', fontSize: 9, bold: false, color: '1A1A1A', align: 'right', margin: 0 });
    slide.addShape(pptx.ShapeType.line, { x: ML, y: BODY_Y + 0.28, w: LEFT_W, h: 0, line: { color: 'E0E0DA', width: 0.5 } });

    // Lignes de prestation : libellé + montant sur une même ligne, taille 9
    const DEV = smcHon.devise || '€';
    let prestY = BODY_Y + 0.32;
    const prestations = (smcHon.prestations || []).filter(p => (parseInt(p.montant) || 0) > 0);
    prestations.forEach(p => {
      slide.addText(p.libelle, {
        x: ML, y: prestY, w: LIB_COL_W, h: 0.28,
        fontFace: 'Segoe UI', fontSize: 9, color: '1A1A1A',
        wrap: false, valign: 'middle', margin: 0
      });
      slide.addText(DEV + ' ' + (parseInt(p.montant) || 0).toLocaleString('fr-FR').replace(/\u202F/g, ' '), {
        x: ML + LIB_COL_W, y: prestY, w: MT_COL_W, h: 0.28,
        fontFace: 'Segoe UI', fontSize: 9, color: '1A1A1A', align: 'right', valign: 'middle', margin: 0
      });
      prestY += 0.32;
    });

    // Clause de réduction (optionnelle)
    if (smcHon.reductionActive && smcHon.reductionMt) {
      slide.addText("En cas de non-réalisation de l'opération envisagée, les honoraires seront réduits à " + DEV + " " + (parseInt(smcHon.reductionMt) || 0).toLocaleString('fr-FR').replace(/\u202F/g, ' ') + " hors taxes.",
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

    // Colonne droite — puces natives pptxgenjs, chaque entrée = paragraphe distinct
    const bullets = [];
    (smcHon.hypotheses || []).forEach(h => {
      bullets.push({
        text: h.txt,
        options: {
          bullet: { code: '25AA' },
          fontSize: 9, color: '404040',
          paraSpaceBefore: 0, paraSpaceAfter: 6,
          lineSpacingMultiple: 1.4
        }
      });
      (h.sub || []).filter(s => s && s.trim()).forEach(s => {
        bullets.push({
          text: s,
          options: {
            bullet: { code: '22A2' },
            indentLevel: 1,
            fontSize: 9, color: '404040',
            paraSpaceBefore: 0, paraSpaceAfter: 6,
            lineSpacingMultiple: 1.4
          }
        });
      });
    });
    slide.addText(bullets, { x: RIGHT_X, y: BODY_Y, w: RIGHT_W, h: BODY_H, fontFace: 'Segoe UI', valign: 'top' });

    // Footer : trait + S-M.C + Page X sur Y (aligné pile sur le trait)
    const footerY = H - MB;
    slide.addShape(pptx.ShapeType.line, { x: ML, y: footerY - 0.12, w: W - ML - MR, h: 0, line: { color: 'E0E0DA', width: 0.5 } });
    slide.addText('S-M.C', { x: ML, y: footerY, w: 2, h: 0.3, fontFace: 'Segoe UI', fontSize: 11, bold: true, color: '1A1A1A', margin: 0 });
    slide.addText('Page ' + (smcHon.pageCur || 1) + ' sur ' + (smcHon.pageTot || 1), {
      x: W - MR - 2, y: footerY, w: 2, h: 0.3,
      fontFace: 'Segoe UI', fontSize: 8, color: '888888', align: 'right', margin: 0
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

// ═══════════════════════════════════════════
// EXPORT PPTX — Slide Calendrier (V8 Calendrier mensuel éditorial)
// ═══════════════════════════════════════════
function smcCalHelpers() {
  const isBiz = d => d.getDay() !== 0 && d.getDay() !== 6;
  const nextBiz = d => { const r = new Date(d); do { r.setDate(r.getDate() + 1); } while (!isBiz(r)); return r; };
  const fmtDuree = dur => {
    const d = parseInt(dur) || 0;
    if (d <= 0) return '0 jour';
    if (d === 1) return '1 jour';
    if (d < 5) return d + ' jours';
    const sem = Math.floor(d / 5), rest = d % 5;
    let out = (sem === 1 ? '1 semaine' : sem + ' semaines');
    if (rest === 1) out += ' et 1 jour';
    else if (rest > 1) out += ' et ' + rest + ' jours';
    return out;
  };
  const mois = ['janv.', 'févr.', 'mars', 'avril', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];
  const fmtPeriod = (s, e) => {
    if (s.getMonth() === e.getMonth() && s.getFullYear() === e.getFullYear())
      return s.getDate() + ' – ' + e.getDate() + ' ' + mois[s.getMonth()] + ' ' + s.getFullYear();
    return s.getDate() + ' ' + mois[s.getMonth()] + ' – ' + e.getDate() + ' ' + mois[e.getMonth()] + ' ' + e.getFullYear();
  };
  const phasePeriods = (dateDebut, phases) => {
    const s0 = new Date(dateDebut); let cur = new Date(s0);
    while (!isBiz(cur)) cur.setDate(cur.getDate() + 1);
    const out = [];
    phases.forEach(ph => {
      const dur = Math.max(1, parseInt(ph.dur) || 1);
      const s = new Date(cur); let e = new Date(cur); let added = 1;
      while (added < dur) { e.setDate(e.getDate() + 1); if (isBiz(e)) added++; }
      out.push({ start: s, end: e });
      cur = nextBiz(e);
    });
    return out;
  };
  const phaseIndex = (day, periods) => {
    if (!isBiz(day)) return -1;
    for (let i = 0; i < periods.length; i++) if (day >= periods[i].start && day <= periods[i].end) return i;
    return -1;
  };
  return { isBiz, nextBiz, fmtDuree, fmtPeriod, phasePeriods, phaseIndex, mois };
}

app.post('/api/export-calendrier-pptx', requireAuth, async (req, res) => {
  try {
    const PptxGenJS = require('pptxgenjs');
    const { smcCal, cibleNom } = req.body;
    if (!smcCal) return res.status(400).json({ success: false, error: 'smcCal manquant.' });

    const H = smcCalHelpers();
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'A4L', width: 11.69, height: 8.27 });
    pptx.layout = 'A4L';
    pptx.title = (cibleNom || 'Propale') + ' — Calendrier prévisionnel';
    pptx.company = 'Saint-Michel Capital';

    const slide = pptx.addSlide();
    slide.background = { color: 'FFFFFF' };

    const W = 11.69, SH = 8.27;
    const ML = 0.7, MR = 0.7, MT = 0.5, MB = 0.4;
    const COL_RIGHT_W = 2.0;
    const TITLE_W = W - ML - MR - COL_RIGHT_W - 0.3;

    // Banner
    slide.addText(
      [
        { text: smcCal.titreBold || 'Calendrier prévisionnel', options: { bold: true, color: '1A1A1A' } },
        { text: ' | ' + (smcCal.titreSuite || ''), options: { bold: false, color: '1A1A1A' } }
      ],
      { x: ML, y: MT, w: TITLE_W, h: 1.1, fontFace: 'Segoe UI', fontSize: 13, align: 'left', valign: 'top', margin: 0, lineSpacingMultiple: 1.4 }
    );

    // Nav droite
    const navX = W - MR - COL_RIGHT_W;
    const navItems = [{ t: 'Contexte', a: false }, { t: 'Périmètre des travaux', a: false }, { t: 'Conditions', a: true }];
    let navY = MT + 0.05;
    navItems.forEach(it => {
      if (it.a) slide.addShape(pptx.ShapeType.rect, { x: navX, y: navY + 0.02, w: 0.03, h: 0.28, fill: { color: '000000' }, line: { color: '000000' } });
      slide.addText(it.t, { x: navX + 0.12, y: navY, w: COL_RIGHT_W - 0.12, h: 0.3, fontFace: 'Segoe UI', fontSize: 8, color: it.a ? '1A1A1A' : '888888', align: 'left', valign: 'top', margin: 0 });
      navY += 0.34;
    });

    // Trait noir
    slide.addShape(pptx.ShapeType.rect, { x: ML, y: MT + 1.2, w: 0.8, h: 0.05, fill: { color: '000000' }, line: { color: '000000' } });

    // Corps : 2 colonnes (calendrier gauche 48%, phases droite 48%)
    const BODY_Y = MT + 1.5;
    const BODY_H = SH - BODY_Y - MB - 0.4;
    const COL_W = (W - ML - MR - 0.4) / 2;
    const LEFT_X = ML;
    const RIGHT_X = ML + COL_W + 0.4;

    const phases = smcCal.phases || [];
    const periods = H.phasePeriods(smcCal.dateDebut || '2026-02-28', phases);

    // ─── Colonne gauche : calendriers mensuels + légende ───
    // Liste des mois à afficher
    const months = [];
    if (periods.length) {
      const s = periods[0].start, e = periods[periods.length - 1].end;
      let y = s.getFullYear(), m = s.getMonth();
      while (y < e.getFullYear() || (y === e.getFullYear() && m <= e.getMonth())) {
        months.push({ year: y, month: m });
        m++; if (m > 11) { m = 0; y++; }
      }
    }
    const twoCol = months.length > 3;
    const monthsPerCol = twoCol ? Math.ceil(months.length / 2) : months.length;
    const mColW = twoCol ? (COL_W - 0.1) / 2 : COL_W;
    const mRowH = 1.15;  // hauteur approximative d'un mois (title + dow + 6 semaines)

    const frMonthNames = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
    const dows = ['LUN', 'MAR', 'MER', 'JEU', 'VEN', 'SAM', 'DIM'];

    months.forEach((mDef, mi) => {
      const col = twoCol ? Math.floor(mi / monthsPerCol) : 0;
      const row = twoCol ? mi % monthsPerCol : mi;
      const mX = LEFT_X + col * (mColW + 0.1);
      const mY = BODY_Y + row * mRowH;
      // Titre mois
      slide.addText(frMonthNames[mDef.month] + ' ' + mDef.year, {
        x: mX, y: mY, w: mColW, h: 0.2,
        fontFace: 'Segoe UI', fontSize: 9, bold: true, color: '1A1A1A', align: 'left', valign: 'top', margin: 0
      });
      // Days of week header
      const cellW = mColW / 7;
      const dowY = mY + 0.22;
      dows.forEach((dw, di) => {
        slide.addText(dw, {
          x: mX + di * cellW, y: dowY, w: cellW, h: 0.12,
          fontFace: 'Segoe UI', fontSize: 5, color: '888888', align: 'center', valign: 'top', margin: 0, charSpacing: 1
        });
      });
      // Ligne sous les dow
      slide.addShape(pptx.ShapeType.line, { x: mX, y: dowY + 0.13, w: mColW, h: 0, line: { color: '1A1A1A', width: 0.5 } });
      // Grille jours
      const first = new Date(mDef.year, mDef.month, 1);
      const daysInMonth = new Date(mDef.year, mDef.month + 1, 0).getDate();
      const firstDow = (first.getDay() + 6) % 7;
      const cellH = 0.14;
      const gridY = dowY + 0.17;
      // Jours hors-mois avant
      for (let i = 0; i < firstDow; i++) {
        const d = new Date(mDef.year, mDef.month, 1 - firstDow + i);
        slide.addText(String(d.getDate()), {
          x: mX + i * cellW, y: gridY, w: cellW, h: cellH,
          fontFace: 'Segoe UI', fontSize: 6, color: 'BBBBBB', align: 'center', valign: 'middle', margin: 0
        });
      }
      // Jours du mois
      for (let day = 1; day <= daysInMonth; day++) {
        const idx = firstDow + day - 1;
        const cx = mX + (idx % 7) * cellW;
        const cy = gridY + Math.floor(idx / 7) * cellH;
        const d = new Date(mDef.year, mDef.month, day);
        const ph = H.phaseIndex(d, periods);
        let color = '404040', fill = null;
        if (ph === 0) {
          // Cercle blanc avec bordure
          slide.addShape(pptx.ShapeType.ellipse, {
            x: cx + cellW / 2 - 0.08, y: cy + cellH / 2 - 0.08, w: 0.16, h: 0.16,
            fill: { color: 'FFFFFF' }, line: { color: '1A1A1A', width: 0.75 }
          });
          color = '1A1A1A';
        } else if (ph === 1) {
          // Rectangle arrondi gris
          slide.addShape(pptx.ShapeType.roundRect, {
            x: cx + 0.01, y: cy + cellH / 2 - 0.06, w: cellW - 0.02, h: 0.12,
            fill: { color: 'E8E8E6' }, line: { color: 'E8E8E6' }, rectRadius: 0.05
          });
          color = '1A1A1A';
        } else if (ph === 2) {
          // Cercle noir plein
          slide.addShape(pptx.ShapeType.ellipse, {
            x: cx + cellW / 2 - 0.08, y: cy + cellH / 2 - 0.08, w: 0.16, h: 0.16,
            fill: { color: '1A1A1A' }, line: { color: '1A1A1A' }
          });
          color = 'FFFFFF';
        } else if (ph === -1) {
          if (!H.isBiz(d)) color = 'BBBBBB';
        }
        slide.addText(String(day), {
          x: cx, y: cy, w: cellW, h: cellH,
          fontFace: 'Segoe UI', fontSize: 6, color: color, align: 'center', valign: 'middle', margin: 0, bold: (ph !== -1)
        });
      }
    });

    // Légende en bas gauche
    const legendY = BODY_Y + BODY_H - 0.6;
    phases.forEach((ph, i) => {
      const legY = legendY + i * 0.18;
      const mark = ['ph1', 'ph2', 'ph3'][i] || 'ph3';
      // Point : cercle blanc bordé, carré gris ou cercle noir
      if (i === 0) slide.addShape(pptx.ShapeType.ellipse, { x: LEFT_X, y: legY + 0.02, w: 0.1, h: 0.1, fill: { color: 'FFFFFF' }, line: { color: '1A1A1A', width: 0.75 } });
      else if (i === 1) slide.addShape(pptx.ShapeType.ellipse, { x: LEFT_X, y: legY + 0.02, w: 0.1, h: 0.1, fill: { color: 'E8E8E6' }, line: { color: 'E8E8E6' } });
      else slide.addShape(pptx.ShapeType.ellipse, { x: LEFT_X, y: legY + 0.02, w: 0.1, h: 0.1, fill: { color: '1A1A1A' }, line: { color: '1A1A1A' } });
      slide.addText((ph.nom || '') + ' (' + H.fmtDuree(ph.dur) + ')', {
        x: LEFT_X + 0.15, y: legY, w: COL_W - 0.2, h: 0.15,
        fontFace: 'Segoe UI', fontSize: 7, color: '404040', align: 'left', valign: 'top', margin: 0
      });
    });

    // ─── Colonne droite : phases détaillées + total ───
    let pY = BODY_Y;
    phases.forEach((ph, i) => {
      const per = periods[i];
      // Trait vertical gauche
      slide.addShape(pptx.ShapeType.line, { x: RIGHT_X, y: pY, w: 0, h: 0.9, line: { color: 'E0E0DA', width: 0.75 } });
      // PHASE N
      slide.addText('PHASE ' + (i + 1), { x: RIGHT_X + 0.12, y: pY, w: COL_W - 0.12, h: 0.15, fontFace: 'Segoe UI', fontSize: 7, color: '888888', charSpacing: 2, margin: 0 });
      // NOM (gros)
      slide.addText((ph.nom || '').toUpperCase(), { x: RIGHT_X + 0.12, y: pY + 0.18, w: COL_W - 0.12, h: 0.3, fontFace: 'Segoe UI', fontSize: 13, bold: true, color: '1A1A1A', margin: 0 });
      // Meta : durée | période
      const dureeTxt = H.fmtDuree(ph.dur).toUpperCase();
      const periodTxt = per ? H.fmtPeriod(per.start, per.end).toUpperCase() : '';
      slide.addText(
        [
          { text: dureeTxt, options: { color: '404040', bold: true } },
          { text: '   |   ', options: { color: 'BBBBBB' } },
          { text: periodTxt, options: { color: '404040', bold: true } }
        ],
        { x: RIGHT_X + 0.12, y: pY + 0.5, w: COL_W - 0.12, h: 0.18, fontFace: 'Segoe UI', fontSize: 7, charSpacing: 1, margin: 0 }
      );
      // Description
      slide.addText(ph.desc || '', { x: RIGHT_X + 0.12, y: pY + 0.7, w: COL_W - 0.12, h: 0.45, fontFace: 'Segoe UI', fontSize: 8, color: '404040', margin: 0, lineSpacingMultiple: 1.3 });
      pY += 1.25;
    });
    // Total
    const totalY = BODY_Y + BODY_H - 0.3;
    slide.addShape(pptx.ShapeType.line, { x: RIGHT_X, y: totalY - 0.05, w: COL_W, h: 0, line: { color: 'E0E0DA', width: 0.5 } });
    slide.addText('DURÉE TOTALE', { x: RIGHT_X, y: totalY, w: COL_W / 2, h: 0.2, fontFace: 'Segoe UI', fontSize: 7, color: '404040', charSpacing: 2, align: 'left', margin: 0 });
    const totalDur = phases.reduce((s, p) => s + Math.max(1, parseInt(p.dur) || 1), 0);
    slide.addText(H.fmtDuree(totalDur).toUpperCase(), { x: RIGHT_X + COL_W / 2, y: totalY, w: COL_W / 2, h: 0.2, fontFace: 'Segoe UI', fontSize: 7, color: '1A1A1A', bold: true, charSpacing: 2, align: 'right', margin: 0 });

    // Pied : trait + S-M.C + Page X sur Y
    const footerY = SH - MB;
    slide.addShape(pptx.ShapeType.line, { x: ML, y: footerY - 0.12, w: W - ML - MR, h: 0, line: { color: 'E0E0DA', width: 0.5 } });
    slide.addText('S-M.C', { x: ML, y: footerY, w: 2, h: 0.3, fontFace: 'Segoe UI', fontSize: 11, bold: true, color: '1A1A1A', margin: 0 });
    slide.addText('Page ' + (smcCal.pageCur || 1) + ' sur ' + (smcCal.pageTot || 1), {
      x: W - MR - 2, y: footerY, w: 2, h: 0.3,
      fontFace: 'Segoe UI', fontSize: 8, color: '888888', align: 'right', margin: 0
    });

    const buf = await pptx.write({ outputType: 'nodebuffer' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="' + (cibleNom || 'Propale').replace(/[^a-zA-Z0-9_-]/g, '_') + '_Calendrier.pptx"');
    res.send(buf);
  } catch (err) {
    console.error('Erreur export Calendrier PPTX:', err.message, err.stack);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.listen(PORT, () => {
  console.log('Serveur Gavroche demarre sur http://localhost:' + PORT);
});
