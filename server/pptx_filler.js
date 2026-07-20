/**
 * pptx_filler.js — Remplissage d'une trame PPTX SANS déformation.
 *
 * Principes absolus :
 *  - JAMAIS modifier ppt/slideLayouts/* ni ppt/slideMasters/*
 *  - Toute nouvelle slide est INSTANCIÉE par duplication d'une slide de la trame
 *  - Le remplissage se fait uniquement sur les balises <a:t> des <p:sp> de la slide
 */

const JSZip = require('jszip');

/* ═══════════════════════════════════════════════════════════════════
 * Helpers XML
 * ═══════════════════════════════════════════════════════════════════ */

function escXml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Remplace le PREMIER <a:t>oldTextExact</a:t> par <a:t>newText</a:t>
function replaceExactText(xml, oldTextExact, newText) {
  const marker = '<a:t>' + oldTextExact + '</a:t>';
  const idx = xml.indexOf(marker);
  if (idx < 0) return { xml, changed: false };
  const replaced = '<a:t>' + escXml(newText) + '</a:t>';
  return { xml: xml.slice(0, idx) + replaced + xml.slice(idx + marker.length), changed: true };
}

// Remplace TOUS les <a:t>oldTextExact</a:t> par <a:t>newText</a:t>
function replaceAllExactText(xml, oldTextExact, newText) {
  let out = xml, changed = 0;
  const marker = '<a:t>' + oldTextExact + '</a:t>';
  const replaced = '<a:t>' + escXml(newText) + '</a:t>';
  let i = 0;
  while ((i = out.indexOf(marker, i)) >= 0) {
    out = out.slice(0, i) + replaced + out.slice(i + marker.length);
    i += replaced.length;
    changed++;
  }
  return { xml: out, changed };
}

/* ═══════════════════════════════════════════════════════════════════
 * PPTXDoc — wrapper JSZip avec accesseurs et duplication de slide
 * ═══════════════════════════════════════════════════════════════════ */

class PPTXDoc {
  constructor(zip) {
    this.zip = zip;
  }

  static async load(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    return new PPTXDoc(zip);
  }

  async save() {
    return await this.zip.generateAsync({ type: 'nodebuffer' });
  }

  // ─── Lecture/écriture des fichiers XML ────────────────────
  async readXml(path) {
    return await this.zip.file(path).async('string');
  }
  writeXml(path, xml) {
    this.zip.file(path, xml);
  }

  async readSlide(n)     { return this.readXml(`ppt/slides/slide${n}.xml`); }
  writeSlide(n, xml)     { this.writeXml(`ppt/slides/slide${n}.xml`, xml); }
  async readSlideRels(n) { return this.readXml(`ppt/slides/_rels/slide${n}.xml.rels`); }
  writeSlideRels(n, xml) { this.writeXml(`ppt/slides/_rels/slide${n}.xml.rels`, xml); }

  // ─── Lister les numéros de slides existantes ────────────
  listSlideNumbers() {
    const nums = [];
    this.zip.forEach((path) => {
      const m = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);
      if (m) nums.push(parseInt(m[1], 10));
    });
    return nums.sort((a, b) => a - b);
  }

  nextSlideNumber() {
    return Math.max(0, ...this.listSlideNumbers()) + 1;
  }

  /**
   * Duplique une slide existante (contenu XML + rels), enregistre les entrées
   * nécessaires (Content_Types, presentation rels, sldIdLst), et retourne le numéro.
   * La nouvelle slide est ajoutée en FIN de sldIdLst (ordre à réorganiser ensuite).
   *
   * @param {number} sourceNum  n° de la slide à cloner (ex 3 = séparateur)
   * @returns {Promise<number>}  n° de la slide nouvellement créée
   */
  async duplicateSlide(sourceNum) {
    const newNum = this.nextSlideNumber();

    // 1) Cloner le XML de la slide et de ses rels
    const srcXml = await this.readSlide(sourceNum);
    const srcRels = await this.readSlideRels(sourceNum);
    this.writeSlide(newNum, srcXml);
    this.writeSlideRels(newNum, srcRels);

    // 2) Ajouter l'Override dans [Content_Types].xml
    let ct = await this.readXml('[Content_Types].xml');
    const overrideTag = `<Override PartName="/ppt/slides/slide${newNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    // Insérer juste avant </Types>
    ct = ct.replace('</Types>', overrideTag + '</Types>');
    this.writeXml('[Content_Types].xml', ct);

    // 3) Ajouter la Relationship dans ppt/_rels/presentation.xml.rels
    let pRels = await this.readXml('ppt/_rels/presentation.xml.rels');
    // Générer un rId unique
    const existingRids = [...pRels.matchAll(/Id="rId(\d+)"/g)].map(m => parseInt(m[1], 10));
    const newRid = 'rId' + (Math.max(0, ...existingRids) + 1);
    const relTag = `<Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newNum}.xml"/>`;
    pRels = pRels.replace('</Relationships>', relTag + '</Relationships>');
    this.writeXml('ppt/_rels/presentation.xml.rels', pRels);

    // 4) Ajouter le sldId dans presentation.xml (en fin de sldIdLst)
    let pres = await this.readXml('ppt/presentation.xml');
    const existingSldIds = [...pres.matchAll(/<p:sldId id="(\d+)"/g)].map(m => parseInt(m[1], 10));
    const newSldId = Math.max(256, ...existingSldIds) + 1;
    const sldIdTag = `<p:sldId id="${newSldId}" r:id="${newRid}"/>`;
    pres = pres.replace('</p:sldIdLst>', sldIdTag + '</p:sldIdLst>');
    this.writeXml('ppt/presentation.xml', pres);

    return newNum;
  }

  /**
   * Supprime une slide (fichiers + Content_Types + presentation rels + sldIdLst).
   */
  async removeSlide(num) {
    // Récupérer le rId de cette slide
    let pRels = await this.readXml('ppt/_rels/presentation.xml.rels');
    const relRe = new RegExp('<Relationship Id="(rId\\d+)"[^>]*Target="slides/slide' + num + '\\.xml"/>', 'g');
    const relMatch = relRe.exec(pRels);
    if (!relMatch) return; // rien à faire
    const rid = relMatch[1];
    // Retirer la relationship
    pRels = pRels.replace(relRe, '');
    this.writeXml('ppt/_rels/presentation.xml.rels', pRels);

    // Retirer le sldId dans presentation.xml
    let pres = await this.readXml('ppt/presentation.xml');
    pres = pres.replace(new RegExp('<p:sldId id="\\d+" r:id="' + rid + '"/>', 'g'), '');
    this.writeXml('ppt/presentation.xml', pres);

    // Retirer l'Override dans Content_Types
    let ct = await this.readXml('[Content_Types].xml');
    ct = ct.replace(new RegExp('<Override PartName="/ppt/slides/slide' + num + '\\.xml"[^/]*/>', 'g'), '');
    this.writeXml('[Content_Types].xml', ct);

    // Supprimer les fichiers
    this.zip.remove('ppt/slides/slide' + num + '.xml');
    this.zip.remove('ppt/slides/_rels/slide' + num + '.xml.rels');
  }

  /**
   * Réordonne les slides selon l'ordre donné (liste de numéros de slide).
   * Ex : reorderSlides([1, 2, 4, 3, 5]) → slide1, slide2, slide4, slide3, slide5.
   */
  async reorderSlides(order) {
    // Récupérer la map rId → slideNum depuis presentation.xml.rels
    const pRels = await this.readXml('ppt/_rels/presentation.xml.rels');
    const ridToNum = {};
    const numToRid = {};
    [...pRels.matchAll(/<Relationship Id="(rId\d+)"[^>]*Target="slides\/slide(\d+)\.xml"\/>/g)].forEach(m => {
      ridToNum[m[1]] = parseInt(m[2], 10);
      numToRid[parseInt(m[2], 10)] = m[1];
    });

    // Extraire les sldId existants (avec leurs ids numériques)
    let pres = await this.readXml('ppt/presentation.xml');
    const sldIdMatches = [...pres.matchAll(/<p:sldId id="(\d+)" r:id="(rId\d+)"\/>/g)];
    const numToSldIdEntry = {};
    sldIdMatches.forEach(m => {
      const rid = m[2];
      const slideNum = ridToNum[rid];
      if (slideNum !== undefined) {
        numToSldIdEntry[slideNum] = `<p:sldId id="${m[1]}" r:id="${rid}"/>`;
      }
    });

    // Construire le nouveau sldIdLst dans l'ordre demandé (les slides non listées sont retirées de l'ordre — mais toujours présentes en fichier)
    const newList = order
      .filter(n => numToSldIdEntry[n])
      .map(n => numToSldIdEntry[n])
      .join('');
    pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst>' + newList + '</p:sldIdLst>');
    this.writeXml('ppt/presentation.xml', pres);
  }

  /**
   * Remplace l'image cover (media/image1.jpg) par un buffer/base64.
   * NB : la Cover référence cette image via <p:bg><a:blipFill><a:blip r:embed="rId2"/></a:blipFill></p:bg>.
   */
  replaceCoverImage(imageBuffer) {
    if (imageBuffer && imageBuffer.length > 0) {
      this.zip.file('ppt/media/image1.jpg', imageBuffer);
    }
  }

  /**
   * Force la Cover (slide1) à un étirement plein cadre (supprime le fillRect avec décalage).
   */
  async normalizeCoverStretch() {
    let s1 = await this.readSlide(1);
    s1 = s1.replace(/<a:stretch>\s*<a:fillRect[^/]*\/>\s*<\/a:stretch>/, '<a:stretch><a:fillRect/></a:stretch>');
    this.writeSlide(1, s1);
  }

  /**
   * Change la référence de layout d'une slide (dans slideN.xml.rels).
   * Ex: changeSlideLayout(9, 17) → slide9 utilise désormais slideLayout17.xml
   */
  async changeSlideLayout(slideNum, newLayoutNum) {
    let rels = await this.readSlideRels(slideNum);
    rels = rels.replace(
      /Target="\.\.\/slideLayouts\/slideLayout\d+\.xml"/,
      `Target="../slideLayouts/slideLayout${newLayoutNum}.xml"`
    );
    this.writeSlideRels(slideNum, rels);
  }

  /**
   * Remplace le contenu texte d'un placeholder identifié par son idx.
   * Le placeholder est identifié par la balise <p:ph ... idx="N" .../> à l'intérieur d'un <p:sp>.
   *
   * @param slideNum n° de slide
   * @param idx  idx du placeholder (par ex. 14, 15, 10)
   * @param paragraphs  tableau de chaînes (une chaîne = un paragraphe/bullet)
   */
  async fillPlaceholderByIdx(slideNum, idx, paragraphs) {
    let xml = await this.readSlide(slideNum);
    // Identifier le <p:sp> qui contient <p:ph ... idx="N" .../>
    const spRegex = /<p:sp>[\s\S]*?<\/p:sp>/g;
    xml = xml.replace(spRegex, (spBlock) => {
      const phRe = new RegExp(`<p:ph [^/]*idx="${idx}"[^/]*/>`);
      if (!phRe.test(spBlock)) return spBlock;
      // Trouvé le bon sp → remplacer son <p:txBody>
      const paras = paragraphs.map(p => {
        return '<a:p><a:r><a:rPr lang="fr-FR" dirty="0"/><a:t>' + escXml(p) + '</a:t></a:r></a:p>';
      }).join('');
      const newTxBody = '<p:txBody><a:bodyPr/><a:lstStyle/>' + paras + '</p:txBody>';
      return spBlock.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTxBody);
    });
    this.writeSlide(slideNum, xml);
  }

  /**
   * Remplace le contenu texte du placeholder title (<p:ph type="title"/>).
   */
  async fillTitle(slideNum, text) {
    let xml = await this.readSlide(slideNum);
    xml = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (spBlock) => {
      if (!/<p:ph type="title"/.test(spBlock)) return spBlock;
      const newTxBody = '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="fr-FR" dirty="0"/><a:t>' + escXml(text) + '</a:t></a:r></a:p></p:txBody>';
      return spBlock.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTxBody);
    });
    this.writeSlide(slideNum, xml);
  }

  /**
   * Ajoute des placeholders vides (idx list) dans une slide (à côté des sp existants),
   * pour supporter tous les emplacements définis par un layout.
   */
  async addPlaceholders(slideNum, idxList) {
    let xml = await this.readSlide(slideNum);
    const existingIdxs = [...xml.matchAll(/<p:ph [^/]*idx="(\d+)"/g)].map(m => parseInt(m[1], 10));
    let nextId = Math.max(4, ...[...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map(m => parseInt(m[1], 10))) + 1;
    const newSps = [];
    for (const idx of idxList) {
      if (existingIdxs.includes(idx)) continue;
      newSps.push(
        '<p:sp><p:nvSpPr><p:cNvPr id="' + nextId + '" name="Text Placeholder ' + idx + '"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
        '<p:nvPr><p:ph type="body" sz="half" idx="' + idx + '"/></p:nvPr>' +
        '</p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="fr-FR"/></a:p></p:txBody></p:sp>'
      );
      nextId++;
    }
    if (newSps.length === 0) { this.writeSlide(slideNum, xml); return; }
    // Injecter juste avant </p:spTree>
    xml = xml.replace('</p:spTree>', newSps.join('') + '</p:spTree>');
    this.writeSlide(slideNum, xml);
  }

  /**
   * Ajoute un placeholder title si absent.
   */
  async addTitlePlaceholder(slideNum) {
    let xml = await this.readSlide(slideNum);
    if (/<p:ph type="title"/.test(xml)) { return; }
    const nextId = Math.max(4, ...[...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map(m => parseInt(m[1], 10))) + 1;
    const titleSp =
      '<p:sp><p:nvSpPr><p:cNvPr id="' + nextId + '" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="title"/></p:nvPr>' +
      '</p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="fr-FR"/></a:p></p:txBody></p:sp>';
    xml = xml.replace('</p:spTree>', titleSp + '</p:spTree>');
    this.writeSlide(slideNum, xml);
  }
}

module.exports = {
  PPTXDoc,
  escXml,
  replaceExactText,
  replaceAllExactText,
};
