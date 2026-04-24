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
// EXPORT PPTX — Slide Calendrier V8 (calendrier mensuel éditorial, refonte 25/04/2026)
// Aligné sur le rendu HTML admin.html
// ═══════════════════════════════════════════
function smcCalHelpersV8() {
  const pad2 = n => String(n).padStart(2, '0');
  const keyOf = d => d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  // Parse ISO "YYYY-MM-DD" en date LOCALE (évite décalage UTC)
  const parseLocal = iso => {
    if (!iso) return null;
    const m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(iso));
    if (!m) return null;
    return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  };
  // Pâques (Meeus/Jones/Butcher)
  const easter = year => {
    const a = year % 19, b = Math.floor(year / 100), c = year % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m2 = Math.floor((a + 11 * h + 22 * l) / 451), month = Math.floor((h + l - 7 * m2 + 114) / 31), day = ((h + l - 7 * m2 + 114) % 31) + 1;
    return new Date(year, month - 1, day);
  };
  const feriesCache = {};
  const feries = year => {
    if (feriesCache[year]) return feriesCache[year];
    const e = easter(year);
    const lp = new Date(e); lp.setDate(e.getDate() + 1);
    const asc = new Date(e); asc.setDate(e.getDate() + 39);
    const pente = new Date(e); pente.setDate(e.getDate() + 50);
    const s = new Set([
      keyOf(new Date(year, 0, 1)), keyOf(lp), keyOf(new Date(year, 4, 1)), keyOf(new Date(year, 4, 8)),
      keyOf(asc), keyOf(pente), keyOf(new Date(year, 6, 14)), keyOf(new Date(year, 7, 15)),
      keyOf(new Date(year, 10, 1)), keyOf(new Date(year, 10, 11)), keyOf(new Date(year, 11, 25))
    ]);
    feriesCache[year] = s;
    return s;
  };
  const isFerie = d => feries(d.getFullYear()).has(keyOf(d));
  const isBiz = d => { const w = d.getDay(); if (w === 0 || w === 6) return false; return !isFerie(d); };
  const nextBiz = d => { const r = new Date(d); do { r.setDate(r.getDate() + 1); } while (!isBiz(r)); return r; };
  const fmtDuree = dur => {
    const d = parseInt(dur) || 0;
    if (d <= 0) return '0 semaine';
    if (d < 5) return '1 semaine';
    const min = Math.floor(d / 5), max = Math.ceil(d / 5);
    if (min === max) return (min === 1 ? '1 semaine' : min + ' semaines');
    return min + '-' + max + ' semaines';
  };
  const mois = ['janv.', 'févr.', 'mars', 'avril', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];
  const fmtPeriod = (s, e) => {
    if (s.getMonth() === e.getMonth() && s.getFullYear() === e.getFullYear())
      return s.getDate() + ' – ' + e.getDate() + ' ' + mois[s.getMonth()] + ' ' + s.getFullYear();
    return s.getDate() + ' ' + mois[s.getMonth()] + ' – ' + e.getDate() + ' ' + mois[e.getMonth()] + ' ' + e.getFullYear();
  };
  const phasePeriods = (dateDebut, phases) => {
    const s0 = parseLocal(dateDebut) || new Date();
    let cur = new Date(s0);
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
  return { isBiz, isFerie, nextBiz, fmtDuree, fmtPeriod, phasePeriods, phaseIndex, parseLocal, keyOf, mois };
}

// Ancien helpers (garde pour compat avec d'autres appels éventuels)
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

    const H = smcCalHelpersV8();
    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'A4L', width: 11.69, height: 8.27 });
    pptx.layout = 'A4L';
    pptx.title = (cibleNom || 'Propale') + ' — Calendrier prévisionnel';
    pptx.company = 'Saint-Michel Capital';

    const slide = pptx.addSlide();
    slide.background = { color: 'FFFFFF' };

    // ─── Dimensions ───
    const W = 11.69, SH = 8.27;
    const ML = 0.55, MR = 0.55, MT = 0.4, MB = 0.25;
    const COL_RIGHT_W = 1.9;
    const TITLE_W = W - ML - MR - COL_RIGHT_W - 0.3;

    // ─── Banner ───
    slide.addText(
      [
        { text: smcCal.titreBold || 'Calendrier prévisionnel', options: { bold: true, color: '1A1A1A' } },
        { text: ' | ' + (smcCal.titreSuite || ''), options: { bold: false, color: '1A1A1A' } }
      ],
      { x: ML, y: MT, w: TITLE_W, h: 0.9, fontFace: 'Segoe UI', fontSize: 13, align: 'left', valign: 'top', margin: 0, lineSpacingMultiple: 1.4 }
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
    slide.addShape(pptx.ShapeType.rect, { x: ML, y: MT + 1.05, w: 0.8, h: 0.05, fill: { color: '000000' }, line: { color: '000000' } });

    // ─── Zones ───
    const BODY_Y = MT + 1.3;
    const FOOTER_Y = SH - MB - 0.25;              // haut du footer S-M.C
    const BOTTOM_Y = FOOTER_Y - 0.5;              // ligne basse (légende + durée totale)
    const BODY_H = BOTTOM_Y - BODY_Y - 0.15;
    const COL_W = (W - ML - MR - 0.3) / 2;
    const LEFT_X = ML;
    const RIGHT_X = ML + COL_W + 0.3;

    const phases = smcCal.phases || [];
    const periods = H.phasePeriods(smcCal.dateDebut || '2026-02-28', phases);

    // ─── Liste des mois à afficher ───
    const months = [];
    if (periods.length) {
      const s = periods[0].start, e = periods[periods.length - 1].end;
      let y = s.getFullYear(), m = s.getMonth();
      while (y < e.getFullYear() || (y === e.getFullYear() && m <= e.getMonth())) {
        months.push({ year: y, month: m });
        m++; if (m > 11) { m = 0; y++; }
      }
    }
    // Mois principal = celui avec le plus de jours de phase
    let mainIdx = 0, bestCount = -1;
    months.forEach((mm, i) => {
      const daysInMonth = new Date(mm.year, mm.month + 1, 0).getDate();
      let count = 0;
      for (let d = 1; d <= daysInMonth; d++) {
        if (H.phaseIndex(new Date(mm.year, mm.month, d), periods) >= 0) count++;
      }
      if (count > bestCount) { bestCount = count; mainIdx = i; }
    });

    // ─── Construction des semaines à afficher pour chaque mois ───
    const frMonthNames = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
    const dows = ['LUN', 'MAR', 'MER', 'JEU', 'VEN', 'SAM', 'DIM'];

    const monthsData = months.map((mDef, mi) => {
      const first = new Date(mDef.year, mDef.month, 1);
      const daysInMonth = new Date(mDef.year, mDef.month + 1, 0).getDate();
      const firstDow = (first.getDay() + 6) % 7;
      const last = new Date(mDef.year, mDef.month, daysInMonth);
      const lastDow = (last.getDay() + 6) % 7;
      const totalCells = firstDow + daysInMonth + (6 - lastDow);
      const cells = [];
      for (let i = 0; i < firstDow; i++) cells.push({ date: new Date(mDef.year, mDef.month, 1 - firstDow + i), inMonth: false });
      for (let d = 1; d <= daysInMonth; d++) cells.push({ date: new Date(mDef.year, mDef.month, d), inMonth: true });
      for (let i = 0; i < (6 - lastDow); i++) cells.push({ date: new Date(mDef.year, mDef.month, daysInMonth + 1 + i), inMonth: false });
      cells.forEach(c => { c.phase = H.phaseIndex(c.date, periods); c.isWeekend = !H.isBiz(c.date); });
      // Découper en semaines de 7
      const weeks = [];
      for (let i = 0; i < cells.length; i += 7) weeks.push(cells.slice(i, i + 7));
      // Si ce n'est pas le mois principal, on ne garde que les semaines avec au moins 1 jour de phase
      const keptWeeks = (mi === mainIdx) ? weeks : weeks.filter(w => w.some(c => c.phase >= 0));
      return { mDef, keptWeeks };
    }).filter(md => md.keptWeeks.length > 0);

    // ─── Dessin des mois ───
    // Hauteur totale du calendrier
    const totalWeeks = monthsData.reduce((s, md) => s + md.keptWeeks.length, 0);
    const MONTH_TITLE_H = 0.2, DOW_H = 0.18, CELL_H = 0.25;
    const availH = BODY_H;
    const neededH = monthsData.length * (MONTH_TITLE_H + DOW_H) + totalWeeks * CELL_H + (monthsData.length - 1) * 0.12;
    const scale = neededH > availH ? (availH / neededH) : 1;
    const titleH = MONTH_TITLE_H * scale;
    const dowH = DOW_H * scale;
    const cellH = CELL_H * scale;
    const mGap = 0.12 * scale;

    const mColW = COL_W;
    const cellW = mColW / 7;
    let curY = BODY_Y;

    monthsData.forEach(({ mDef, keptWeeks }) => {
      const mX = LEFT_X;
      const mY = curY;
      // Titre mois
      slide.addText(frMonthNames[mDef.month] + ' ' + mDef.year, {
        x: mX, y: mY, w: mColW, h: titleH,
        fontFace: 'Segoe UI', fontSize: 11, bold: true, color: '1A1A1A', align: 'left', valign: 'top', margin: 0
      });
      // DOW + trait gris (raccourci à droite)
      const dowY = mY + titleH;
      dows.forEach((dw, di) => {
        slide.addText(dw, {
          x: mX + di * cellW, y: dowY, w: cellW, h: dowH,
          fontFace: 'Segoe UI', fontSize: 6, color: '888888', align: 'center', valign: 'top', margin: 0, charSpacing: 1.5
        });
      });
      // Trait gris sous les DOW (raccourci à droite de ~0.2")
      slide.addShape(pptx.ShapeType.line, { x: mX, y: dowY + dowH, w: mColW - 0.2, h: 0, line: { color: 'D5D5D2', width: 0.5 } });
      // Semaines
      const gridY = dowY + dowH + 0.04;
      keptWeeks.forEach((week, wi) => {
        const rowY = gridY + wi * cellH;
        week.forEach((c, ci) => {
          const cx = mX + ci * cellW;
          const cy = rowY;
          const d = c.date;
          const ph = c.phase;
          let textColor = '1A1A1A', bold = true;
          if (ph === 0) {
            // Cercle blanc bordé
            const circleD = Math.min(cellW * 0.85, cellH * 0.85);
            slide.addShape(pptx.ShapeType.ellipse, {
              x: cx + (cellW - circleD) / 2, y: cy + (cellH - circleD) / 2,
              w: circleD, h: circleD,
              fill: { color: 'FFFFFF' }, line: { color: '1A1A1A', width: 0.75 }
            });
          } else if (ph === 1) {
            // Pilule grise : rectangle arrondi selon start/mid/end
            const prev = ci > 0 ? week[ci - 1] : null;
            const next = ci < 6 ? week[ci + 1] : null;
            const hasPrev = prev && prev.phase === 1;
            const hasNext = next && next.phase === 1;
            const barH = cellH * 0.7;
            const barY = cy + (cellH - barH) / 2;
            let rx = cx, rw = cellW;
            if (!hasPrev) rx = cx + 0.02;
            if (!hasNext) rw = cellW - 0.02 - (!hasPrev ? 0.02 : 0);
            else if (!hasPrev) rw = cellW - 0.02;
            const radius = !hasPrev && !hasNext ? barH / 2 : (!hasPrev || !hasNext ? barH / 2 : 0);
            slide.addShape(pptx.ShapeType.roundRect, {
              x: rx, y: barY, w: rw, h: barH,
              fill: { color: 'F0F0F0' }, line: { color: 'F0F0F0' }, rectRadius: radius
            });
          } else if (ph === 2) {
            // Cercle noir plein
            const circleD = Math.min(cellW * 0.85, cellH * 0.85);
            slide.addShape(pptx.ShapeType.ellipse, {
              x: cx + (cellW - circleD) / 2, y: cy + (cellH - circleD) / 2,
              w: circleD, h: circleD,
              fill: { color: '1A1A1A' }, line: { color: '1A1A1A' }
            });
            textColor = 'FFFFFF';
          } else {
            // Hors phase : jour ouvré = noir, weekend/hors mois = gris clair
            if (!c.inMonth || c.isWeekend) { textColor = 'C8C8C8'; bold = false; }
          }
          slide.addText(String(d.getDate()), {
            x: cx, y: cy, w: cellW, h: cellH,
            fontFace: 'Segoe UI', fontSize: 9, color: textColor, align: 'center', valign: 'middle', margin: 0, bold: bold
          });
        });
      });
      curY += titleH + dowH + 0.04 + keptWeeks.length * cellH + mGap;
    });

    // ─── Colonne droite : 3 phases (titre + meta + bullets) ───
    const PHASE_H = (BODY_H - 0.2) / phases.length;
    phases.forEach((ph, i) => {
      const per = periods[i];
      const pY = BODY_Y + i * PHASE_H + 0.1;
      // Trait vertical gauche
      slide.addShape(pptx.ShapeType.line, { x: RIGHT_X, y: pY, w: 0, h: PHASE_H - 0.2, line: { color: 'E0E0DA', width: 0.75 } });
      // PHASE N (petit caps)
      slide.addText('PHASE ' + (i + 1), {
        x: RIGHT_X + 0.15, y: pY, w: COL_W - 0.15, h: 0.15,
        fontFace: 'Segoe UI', fontSize: 7, color: '888888', charSpacing: 2, margin: 0, bold: false
      });
      // NOM (gros)
      slide.addText((ph.nom || '').toUpperCase(), {
        x: RIGHT_X + 0.15, y: pY + 0.2, w: COL_W - 0.15, h: 0.35,
        fontFace: 'Segoe UI', fontSize: 14, bold: true, color: '1A1A1A', margin: 0, charSpacing: 0.3
      });
      // Meta : durée | période
      const dureeTxt = H.fmtDuree(ph.dur).toUpperCase();
      const periodTxt = per ? H.fmtPeriod(per.start, per.end).toUpperCase() : '';
      slide.addText(
        [
          { text: dureeTxt, options: { color: '404040', bold: true } },
          { text: '   |   ', options: { color: 'BBBBBB' } },
          { text: periodTxt, options: { color: '404040', bold: true } }
        ],
        { x: RIGHT_X + 0.15, y: pY + 0.6, w: COL_W - 0.15, h: 0.18, fontFace: 'Segoe UI', fontSize: 7, charSpacing: 1.5, margin: 0 }
      );
      // Bullets Retinax
      const pts = (ph.points || []);
      if (pts.length) {
        const bullets = pts.map(pt => ({
          text: pt,
          options: { bullet: { code: '25AA' }, fontSize: 8, color: '404040', paraSpaceBefore: 0, paraSpaceAfter: 4, lineSpacingMultiple: 1.35 }
        }));
        slide.addText(bullets, {
          x: RIGHT_X + 0.15, y: pY + 0.85, w: COL_W - 0.15, h: PHASE_H - 1,
          fontFace: 'Segoe UI', valign: 'top'
        });
      }
    });

    // ─── Ligne basse : légende à gauche + durée totale à droite ───
    // Légende en ligne : 3 items côte à côte
    let legX = ML;
    phases.forEach((ph, i) => {
      const circleY = BOTTOM_Y + 0.05;
      const circleD = 0.12;
      if (i === 0) slide.addShape(pptx.ShapeType.ellipse, { x: legX, y: circleY, w: circleD, h: circleD, fill: { color: 'FFFFFF' }, line: { color: '1A1A1A', width: 0.75 } });
      else if (i === 1) slide.addShape(pptx.ShapeType.ellipse, { x: legX, y: circleY, w: circleD, h: circleD, fill: { color: 'F0F0F0' }, line: { color: 'F0F0F0' } });
      else slide.addShape(pptx.ShapeType.ellipse, { x: legX, y: circleY, w: circleD, h: circleD, fill: { color: '1A1A1A' }, line: { color: '1A1A1A' } });
      const legTxt = (ph.nom || '') + ' (' + H.fmtDuree(ph.dur) + ')';
      // Estimation largeur texte : ~0.07 par caractère à font 7
      const approxW = legTxt.length * 0.055 + 0.25;
      slide.addText(legTxt, {
        x: legX + circleD + 0.08, y: BOTTOM_Y, w: approxW, h: 0.3,
        fontFace: 'Segoe UI', fontSize: 8, color: '404040', align: 'left', valign: 'middle', margin: 0
      });
      legX += circleD + 0.08 + approxW + 0.2;
    });
    // Durée totale (tout à droite)
    const totalDur = phases.reduce((s, p) => s + Math.max(1, parseInt(p.dur) || 1), 0);
    const totalValue = H.fmtDuree(totalDur).toUpperCase();
    slide.addText('DURÉE TOTALE', {
      x: W - MR - 2.7, y: BOTTOM_Y, w: 1.4, h: 0.3,
      fontFace: 'Segoe UI', fontSize: 8, color: '404040', charSpacing: 1.5, align: 'right', valign: 'middle', margin: 0
    });
    slide.addText(totalValue, {
      x: W - MR - 1.2, y: BOTTOM_Y, w: 1.2, h: 0.3,
      fontFace: 'Segoe UI', fontSize: 8, color: '1A1A1A', bold: true, charSpacing: 1.5, align: 'right', valign: 'middle', margin: 0
    });

    // ─── Footer S-M.C ───
    slide.addShape(pptx.ShapeType.line, { x: ML, y: FOOTER_Y, w: W - ML - MR, h: 0, line: { color: 'E0E0DA', width: 0.5 } });
    slide.addText('S-M.C', { x: ML, y: FOOTER_Y + 0.08, w: 2, h: 0.25, fontFace: 'Segoe UI', fontSize: 11, bold: true, color: '1A1A1A', margin: 0 });
    slide.addText('Page ' + (smcCal.pageCur || 1) + ' sur ' + (smcCal.pageTot || 1), {
      x: W - MR - 2, y: FOOTER_Y + 0.08, w: 2, h: 0.25,
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
