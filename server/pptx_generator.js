/**
 * pptx_generator.js — Orchestration : à partir d'une spec, produit un PPTX
 * en s'appuyant sur la trame INTACTE (pptx_filler.js).
 */

const fs = require('fs');
const path = require('path');
const { PPTXDoc, escXml } = require('./pptx_filler');

/* ═══════════════════════════════════════════════════════════════════
 * Numéros de slide dans la trame originale (référence stable)
 * ═══════════════════════════════════════════════════════════════════ */
const SRC = {
  COVER:       1,   // layout 1 (Cover)
  SOMMAIRE:    2,   // layout 20 (Blank + textes lettre/sommaire)
  SEP_1:       3,   // layout 2 (Section Bach. — "1 Contexte de l'opération")
  SEP_2:       4,   // layout 2 (Section Bach. — "2 Périmètre des travaux")
  SEP_3:       5,   // layout 2 (Section Bach. — "3 Conditions d'intervention")
  CALENDRIER:  6,   // layout 18 (Calendrier vierge — gabarit)
  BLANK:       7,   // layout 20 (vide — utilitaire)
  BACKCOVER:   8    // layout 1 (backcover)
};

const LAYOUT = {
  COVER: 1,
  SECTION: 2,
  SCOPE: 17,
  CALENDRIER: 18,
  BUDGET: 19,
  BLANK: 20
};

/* ═══════════════════════════════════════════════════════════════════
 * Slide 1 : Cover — remplacer titre + image de fond
 * ═══════════════════════════════════════════════════════════════════ */
async function fillCover(pptx, spec) {
  const titre = spec.titreCover || `Rapport de due diligence financière — ${spec.cibleNom || ''}`.trim();

  // Remplacer le titre
  let xml = await pptx.readSlide(SRC.COVER);
  xml = xml.replace(
    /<a:t>Rapport de due diligence financière<\/a:t>/,
    '<a:t>' + escXml(titre) + '</a:t>'
  );
  // Étirement plein cadre
  xml = xml.replace(
    /<a:stretch>\s*<a:fillRect[^/]*\/>\s*<\/a:stretch>/,
    '<a:stretch><a:fillRect/></a:stretch>'
  );
  pptx.writeSlide(SRC.COVER, xml);

  // Remplacer l'image cover si fournie (base64 dataURL ou raw base64)
  if (spec.coverImageBase64) {
    const b64 = spec.coverImageBase64.replace(/^data:image\/[^;]+;base64,/, '');
    pptx.replaceCoverImage(Buffer.from(b64, 'base64'));
  }
}

/* ═══════════════════════════════════════════════════════════════════
 * Slide 2 : Lettre + Sommaire — remplacer les items du sommaire
 * ═══════════════════════════════════════════════════════════════════ */
async function fillSommaire(pptx, spec, sectionsOrdered) {
  // sectionsOrdered = liste ordonnée des sections cochées :
  // [{ num: '1', titre: 'Contexte de l'opération', page: 3 }, ...]

  let xml = await pptx.readSlide(SRC.SOMMAIRE);

  // Remplacer les 3 items par défaut (Contexte / Périmètre / Conditions).
  // On procède par identification des <a:t>...</a:t> qui portent les titres actuels
  // et on remplace en respectant l'ordre.
  const defaultItems = [
    'Contexte de l’opération',
    'Périmètre des travaux…………………………………………………………………………………',
    'Conditions d’intervention'
  ];
  // Pour la version simple : on remplace juste les libellés des 3 items visibles.
  // (Les items 4/5 pour Approche/Références seront ajoutés dans une itération future.)
  for (let i = 0; i < Math.min(sectionsOrdered.length, 3); i++) {
    xml = xml.replace(
      new RegExp('<a:t>' + defaultItems[i].replace(/[\-\/\\^$*+?.()|[\]{}]/g, '\\$&') + '<\/a:t>'),
      '<a:t>' + escXml(sectionsOrdered[i].titre) + '</a:t>'
    );
  }
  pptx.writeSlide(SRC.SOMMAIRE, xml);
}

/* ═══════════════════════════════════════════════════════════════════
 * Séparateur (Section Bach.) — remplacer numéro et titre
 * ═══════════════════════════════════════════════════════════════════ */
async function fillSeparator(pptx, slideNum, oldNum, oldTitre, newNum, newTitre) {
  let xml = await pptx.readSlide(slideNum);
  // Le numéro est dans un <a:t>N</a:t> (donné par oldNum)
  xml = xml.replace(new RegExp('<a:t>' + oldNum + '<\/a:t>'), '<a:t>' + escXml(newNum) + '</a:t>');
  xml = xml.replace(
    new RegExp('<a:t>' + oldTitre.replace(/[\-\/\\^$*+?.()|[\]{}]/g, '\\$&') + '<\/a:t>'),
    '<a:t>' + escXml(newTitre) + '</a:t>'
  );
  pptx.writeSlide(slideNum, xml);
}

/* ═══════════════════════════════════════════════════════════════════
 * Slide Scope of work — créée par duplication du gabarit (slide 6)
 * puis reroutée vers le layout 17
 * ═══════════════════════════════════════════════════════════════════ */
async function createScopeOfWork(pptx, obj, page, total) {
  // 1) Duplique le gabarit vierge (slide 6)
  const newNum = await pptx.duplicateSlide(SRC.CALENDRIER);
  // 2) Reroute vers layout 17 (Scope of work)
  await pptx.changeSlideLayout(newNum, LAYOUT.SCOPE);
  // 3) Ajoute placeholder title + idx=2 (petit sous-titre)
  await pptx.addTitlePlaceholder(newNum);
  await pptx.addPlaceholders(newNum, [2]);

  // Titre nettoyé (retirer [le cas échéant])
  const titreClean = (obj.titre || '').replace(/\s*\[le cas échéant\]\s*/gi, '').trim();

  // Mapping selon la version corrigée par François :
  //  - Title (haut-gauche)  : titre court (à défaut : titreClean)
  //  - idx=2 (sous-titre)   : titre long
  //  - idx=14 (col gauche)  : la FINALITÉ entre guillemets
  //  - idx=10 (bandeau haut): "Diligences liées à l'objectif « titre »"
  //  - idx=15 (col droite)  : diligences en bullets
  await pptx.fillTitle(newNum, obj.titreCourt || titreClean);
  await pptx.fillPlaceholderByIdx(newNum, 2, [titreClean]);
  if (obj.finalite) {
    await pptx.fillPlaceholderByIdx(newNum, 14, [`« ${obj.finalite} »`]);
  } else {
    await pptx.fillPlaceholderByIdx(newNum, 14, [' ']);
  }
  await pptx.fillPlaceholderByIdx(newNum, 10, [`Diligences liées à l'objectif « ${titreClean} »`]);
  // Diligences en bullets : chaque diligence niveau 0 (bullet carré ▪),
  // chaque sub niveau 1 (bullet tiret –). Line spacing 1.4 partout.
  const LS = 140; // 1.4
  const dilBullets = [];
  (obj.diligences || []).forEach(d => {
    if (typeof d === 'string') {
      if (d.trim()) dilBullets.push({ text: d, level: 0, bullet: '▪', lineSpacing: LS });
      return;
    }
    const t = d.text || '';
    if (t.trim()) dilBullets.push({ text: t, level: 0, bullet: '▪', lineSpacing: LS });
    const subs = d.subs || d.sub || [];
    subs.forEach(s => {
      if (s && String(s).trim()) dilBullets.push({ text: String(s), level: 1, bullet: '–', lineSpacing: LS });
    });
  });
  await pptx.fillPlaceholderByIdx(newNum, 15, dilBullets.length ? dilBullets : [' ']);

  return newNum;
}

/* ═══════════════════════════════════════════════════════════════════
 * Slide Calendrier — remplit le gabarit slide 6 (layout 18) déjà en place
 * ═══════════════════════════════════════════════════════════════════ */
async function fillCalendrier(pptx, cal) {
  const slideNum = SRC.CALENDRIER;
  // Titre en haut à droite (idx=10) : "Calendrier prévisionnel"
  await pptx.fillPlaceholderByIdx(slideNum, 10, ['Calendrier prévisionnel']);

  // Col gauche (idx=14) : liste des phases
  const phases = (cal && cal.phases) || [];
  const leftLines = phases.map((p, i) => {
    const nom = p.nom || `Phase ${i + 1}`;
    const dates = [p.debut, p.fin].filter(Boolean).join(' → ');
    return dates ? `${nom} : ${dates}` : nom;
  });
  await pptx.fillPlaceholderByIdx(slideNum, 14, leftLines.length ? leftLines : [' ']);

  // Col droite (idx=15) : livrables
  const livr = (cal && cal.livrables) || [];
  const rightLines = livr.length ? livr : [' '];
  await pptx.fillPlaceholderByIdx(slideNum, 15, rightLines);
}

/* ═══════════════════════════════════════════════════════════════════
 * Slide Budget — créée par duplication du gabarit + reroute layout 19
 * ═══════════════════════════════════════════════════════════════════ */
async function createBudget(pptx, hon) {
  const newNum = await pptx.duplicateSlide(SRC.CALENDRIER);
  await pptx.changeSlideLayout(newNum, LAYOUT.BUDGET);
  await pptx.addTitlePlaceholder(newNum);
  await pptx.addPlaceholders(newNum, [2, 16, 17, 18, 19]);

  const devise = hon.devise || '€';
  const prestations = (hon.prestations || []).filter(p => (parseInt(p.montant) || 0) > 0);

  // Mapping selon la version corrigée par François :
  //  - Title             : "Honoraires"
  //  - idx=10 (bandeau)  : 2 paragraphes → "Honoraires" + " | Notre équipe…"
  //  - idx=14 (col gauche): libellé(s) de prestation(s)
  //  - idx=16            : montant(s) HT correspondant(s)
  //  - idx=17            : LIBELLÉ "Total HT" (pas le montant !)
  //  - idx=18            : clause de réduction
  //  - idx=19            : durée + livrIntro + livrables (SANS les hypothèses)
  //  - idx=15 (col droite): HYPOTHÈSES complètes

  await pptx.fillTitle(newNum, 'Honoraires');

  await pptx.fillPlaceholderByIdx(newNum, 10, [
    'Honoraires',
    ' | ' + (hon.titreSuite || 'Notre équipe sera mobilisée sur toute la période d’exécution des travaux.')
  ]);

  // Prestations : libellés + montants alignés
  const libs = prestations.map(p => p.libelle || '');
  const mts  = prestations.map(p => devise + ' ' + (parseInt(p.montant) || 0).toLocaleString('fr-FR'));
  await pptx.fillPlaceholderByIdx(newNum, 14, libs.length ? libs : [' ']);
  await pptx.fillPlaceholderByIdx(newNum, 16, mts.length ? mts : [' ']);

  // idx=17 : LIBELLÉ Total HT (pas le montant)
  await pptx.fillPlaceholderByIdx(newNum, 17, ['Total HT']);

  // idx=18 : clause réduction
  if (hon.reductionActive && hon.reductionMt) {
    await pptx.fillPlaceholderByIdx(newNum, 18, [
      'En cas de non-réalisation de l’opération envisagée, les honoraires seront réduits à ' +
      devise + ' ' + parseInt(hon.reductionMt).toLocaleString('fr-FR') + ' hors taxes.'
    ]);
  }

  // idx=19 : durée + livrables uniquement
  const dureeLivr = [];
  if (hon.dureeTxt)  dureeLivr.push(hon.dureeTxt);
  if (hon.livrIntro) dureeLivr.push(hon.livrIntro);
  (hon.livrables || []).forEach(l => dureeLivr.push('— ' + l));
  await pptx.fillPlaceholderByIdx(newNum, 19, dureeLivr.length ? dureeLivr : [' ']);

  // idx=15 : hypothèses complètes (chaque txt + chaque sub sur son propre paragraphe)
  const hyps = [];
  (hon.hypotheses || []).forEach(h => {
    if (h.txt) hyps.push(h.txt);
    (h.sub || []).forEach(s => hyps.push(s));
  });
  await pptx.fillPlaceholderByIdx(newNum, 15, hyps.length ? hyps : [' ']);

  return newNum;
}

/* ═══════════════════════════════════════════════════════════════════
 * ORCHESTRATEUR PRINCIPAL
 * ═══════════════════════════════════════════════════════════════════ */

/**
 * spec = {
 *   trameChoice: 'arial'|'segoe',
 *   coverImageBase64: '...',
 *   titreCover: '...',
 *   cibleNom: 'ABC SA',
 *   projetNom: 'PROJET GAVROCHE',
 *   parties: { p1:true, p2:true, p3:true, p4:true, p5:false }, // Contexte/Approche/Périmètre/Conditions/Références
 *   perimetre: [
 *     { id, titre, finalite, diligences:[{text, subs:[]}, ...] }, ...
 *   ], // uniquement les objectifs COCHÉS
 *   calendrier: { phases:[{nom,debut,fin}], livrables:[] },
 *   honoraires: { devise, prestations:[{libelle, montant}], reductionActive, reductionMt, dureeTxt, livrIntro, livrables:[], hypotheses:[{txt,sub:[]}] }
 * }
 */
async function generatePropale(trameBuf, spec) {
  const pptx = await PPTXDoc.load(trameBuf);

  // ─── 1) Cover ────────────────────────────────────────────
  await fillCover(pptx, spec);

  // ─── 2) Détermine l'ordre des sections cochées ──────────
  const partiesSpec = [
    { key: 'p1', num: '1', titre: 'Contexte de l’opération',        srcSlide: SRC.SEP_1 },
    { key: 'p2', num: '2', titre: 'Approche',                             srcSlide: null       },
    { key: 'p3', num: '3', titre: 'Périmètre des travaux',                srcSlide: SRC.SEP_2  },
    { key: 'p4', num: '4', titre: 'Conditions d’intervention',        srcSlide: SRC.SEP_3  },
    { key: 'p5', num: '5', titre: 'Références',                           srcSlide: null       }
  ];
  const partiesCochees = partiesSpec.filter(p => (spec.parties || {})[p.key]);
  // Renumérote 1..N les sections cochées
  partiesCochees.forEach((p, i) => { p.numFinal = String(i + 1); });

  // ─── 3) Séparateurs pour chaque section cochée ─────────
  const separatorSlideNums = {}; // key → n° de slide
  for (const p of partiesCochees) {
    let slideNum;
    if (p.srcSlide) {
      // On récupère la slide existante et on modifie numéro + titre
      slideNum = p.srcSlide;
      await fillSeparator(
        pptx, slideNum,
        p.srcSlide === SRC.SEP_1 ? '1' : (p.srcSlide === SRC.SEP_2 ? '2' : '3'),
        p.srcSlide === SRC.SEP_1 ? 'Contexte de l’opération'
          : (p.srcSlide === SRC.SEP_2 ? 'Périmètre des travaux' : 'Conditions d’intervention'),
        p.numFinal, p.titre
      );
    } else {
      // Approche ou Références → duplique le séparateur 1
      slideNum = await pptx.duplicateSlide(SRC.SEP_1);
      await fillSeparator(pptx, slideNum, '1', 'Contexte de l’opération', p.numFinal, p.titre);
    }
    separatorSlideNums[p.key] = slideNum;
  }

  // ─── 4) Slides Scope of work (1 par objectif coché) ────
  const scopeSlideNums = [];
  if (partiesCochees.find(p => p.key === 'p3')) {
    const perimList = spec.perimetre || [];
    const total = perimList.length;
    for (let i = 0; i < perimList.length; i++) {
      const num = await createScopeOfWork(pptx, perimList[i], i + 1, total);
      scopeSlideNums.push(num);
    }
  }

  // ─── 5) Calendrier + Budget (si Conditions cochée) ─────
  let calendrierNum = null, budgetNum = null;
  if (partiesCochees.find(p => p.key === 'p4')) {
    calendrierNum = SRC.CALENDRIER;
    await fillCalendrier(pptx, spec.calendrier || {});
    budgetNum = await createBudget(pptx, spec.honoraires || {});
  }

  // ─── 6) Sommaire ────────────────────────────────────────
  await fillSommaire(pptx, spec, partiesCochees.map(p => ({ num: p.numFinal, titre: p.titre })));

  // ─── 7) Assemble l'ordre final ─────────────────────────
  const order = [SRC.COVER, SRC.SOMMAIRE];
  for (const p of partiesCochees) {
    order.push(separatorSlideNums[p.key]);
    if (p.key === 'p3') {
      order.push(...scopeSlideNums);
    } else if (p.key === 'p4') {
      if (calendrierNum) order.push(calendrierNum);
      if (budgetNum) order.push(budgetNum);
    }
  }
  order.push(SRC.BACKCOVER);

  // Supprimer la slide 7 (blank vestige) et toute slide non listée
  await pptx.removeSlide(SRC.BLANK);
  await pptx.reorderSlides(order);

  return await pptx.save();
}

module.exports = { generatePropale, LAYOUT, SRC };
