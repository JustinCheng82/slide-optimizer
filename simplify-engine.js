/* Lucid Slides — read-only PPTX analysis engine.
 *
 * Safety contract:
 * - Reads ZIP entries without replacing or serializing any OOXML part.
 * - Never generates a PPTX, Blob, or modified package.
 * - Never applies threshold-based text, style, image, or structure edits.
 * - Produces review findings and validates optional AI proposals only.
 */

const SLIDE_PATH = /^ppt\/slides\/slide(\d+)\.xml$/;
const REQUIRED_PARTS = ["[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"];
const PLACEHOLDER_TEXT = "AI analysis required — Lucid Slides will not guess which wording or element is important.";

function decodeXml(value = "") {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([\da-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16)))
    .replace(/&amp;/g, "&");
}

function readAttribute(xml = "", name) {
  const match = xml.match(new RegExp(`\\b${name}="([^"]*)"`));
  return match ? decodeXml(match[1]) : "";
}

function textFromXml(xml = "") {
  return Array.from(xml.matchAll(/<a:t\b[^>]*>([\s\S]*?)<\/a:t>/g), (match) => decodeXml(match[1])).join("");
}

function wordCount(text = "") {
  const value = String(text).trim();
  return value ? value.split(/\s+/).length : 0;
}

function boldWordCount(paragraphXml = "") {
  let count = 0;
  for (const run of paragraphXml.matchAll(/<a:r\b[\s\S]*?<\/a:r>/g)) {
    const runXml = run[0];
    const properties = runXml.match(/<a:rPr\b[^>]*>/)?.[0] || "";
    if (/\bb="(?:1|true)"/.test(properties)) count += wordCount(textFromXml(runXml));
  }
  return count;
}

function elementType(elementXml, tagName) {
  if (tagName === "pic") return "image";
  if (/<a:tbl\b/.test(elementXml)) return "table";
  if (/<c:chart\b/.test(elementXml)) return "chart";
  const placeholder = elementXml.match(/<p:ph\b[^>]*>/)?.[0] || "";
  const placeholderType = readAttribute(placeholder, "type");
  if (placeholderType === "title" || placeholderType === "ctrTitle") return "title";
  return "text";
}

export function analyzeSlideXml(xml, slideNumber) {
  if (typeof xml !== "string" || !/<p:sld\b/.test(xml)) {
    throw new Error(`Slide ${slideNumber} is not valid PresentationML.`);
  }

  const elements = [];
  let fallbackId = 0;
  for (const match of xml.matchAll(/<p:(sp|graphicFrame|pic)\b[\s\S]*?<\/p:\1>/g)) {
    fallbackId += 1;
    const tagName = match[1];
    const elementXml = match[0];
    const nonVisual = elementXml.match(/<p:cNvPr\b[^>]*>/)?.[0] || "";
    const objectId = readAttribute(nonVisual, "id") || `unknown-${fallbackId}`;
    const name = readAttribute(nonVisual, "name") || `${tagName} ${objectId}`;
    const paragraphs = [];
    let paragraphIndex = 0;
    for (const paragraphMatch of elementXml.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)) {
      const paragraphXml = paragraphMatch[0];
      const text = textFromXml(paragraphXml);
      if (!text.trim()) continue;
      const words = wordCount(text);
      paragraphs.push({
        index: paragraphIndex,
        text,
        wordCount: words,
        boldWordCount: boldWordCount(paragraphXml),
        hasHyperlink: /<a:hlinkClick\b/.test(paragraphXml),
        isBullet: /<a:bu(?:Char|AutoNum|Blip)\b/.test(paragraphXml),
      });
      paragraphIndex += 1;
    }
    const relationshipId = readAttribute(elementXml.match(/<a:blip\b[^>]*>/)?.[0] || "", "r:embed") ||
      readAttribute(elementXml.match(/<a:blip\b[^>]*>/)?.[0] || "", "r:link");
    elements.push({
      objectId,
      name,
      type: elementType(elementXml, tagName),
      text: paragraphs.map((paragraph) => paragraph.text).join("\n"),
      paragraphs,
      relationshipId: relationshipId || null,
      hasHyperlink: paragraphs.some((paragraph) => paragraph.hasHyperlink),
    });
  }

  return {
    slide: Number(slideNumber),
    elements,
    counts: {
      images: elements.filter((element) => element.type === "image").length,
      tables: elements.filter((element) => element.type === "table").length,
      charts: elements.filter((element) => element.type === "chart").length,
      hyperlinks: elements.reduce(
        (sum, element) => sum + element.paragraphs.filter((paragraph) => paragraph.hasHyperlink).length,
        0,
      ),
      words: elements.reduce((sum, element) => sum + wordCount(element.text), 0),
    },
  };
}

function findingId(slide, element, paragraph, rule) {
  return `slide-${slide}-element-${element.objectId}-paragraph-${paragraph?.index ?? "all"}-rule-${rule}`;
}

export function buildLocalFindings(slides) {
  const findings = [];
  for (const slide of slides) {
    for (const element of slide.elements) {
      if (element.type === "title" && wordCount(element.text) > 10) {
        findings.push({
          id: findingId(slide.slide, element, null, 2),
          slide: slide.slide,
          objectId: element.objectId,
          elementName: element.name,
          elementType: element.type,
          originalText: element.text,
          proposedText: null,
          placeholderText: PLACEHOLDER_TEXT,
          rule: 2,
          explanation: "The title exceeds the rule-of-thumb length. Shortening it requires understanding the slide's actual takeaway, so no rewrite was attempted.",
          actionable: false,
          source: "local-analysis",
        });
      }

      for (const paragraph of element.paragraphs) {
        if (element.type !== "title" && paragraph.wordCount > 12) {
          findings.push({
            id: findingId(slide.slide, element, paragraph, 3),
            slide: slide.slide,
            objectId: element.objectId,
            elementName: element.name,
            elementType: element.type,
            originalText: paragraph.text,
            proposedText: null,
            placeholderText: PLACEHOLDER_TEXT,
            rule: 3,
            explanation: "This passage may be difficult to scan. Deciding what is repeated or nonessential is a meaning-based judgment, so every word was preserved.",
            actionable: false,
            source: "local-analysis",
          });
        }
        if (paragraph.wordCount > 0 && paragraph.boldWordCount / paragraph.wordCount > 0.2) {
          findings.push({
            id: findingId(slide.slide, element, paragraph, 4),
            slide: slide.slide,
            objectId: element.objectId,
            elementName: element.name,
            elementType: element.type,
            originalText: paragraph.text,
            proposedText: null,
            placeholderText: PLACEHOLDER_TEXT,
            rule: 4,
            explanation: "More than about 20% of this statement is bold. Choosing which words deserve emphasis requires semantic judgment, so formatting was not changed.",
            actionable: false,
            source: "local-analysis",
          });
        }
      }
    }

    if (slide.counts.charts) {
      findings.push({
        id: `slide-${slide.slide}-chart-rule-8`,
        slide: slide.slide,
        objectId: null,
        elementName: "Chart",
        elementType: "chart",
        originalText: "Chart content and styling preserved exactly.",
        proposedText: null,
        placeholderText: PLACEHOLDER_TEXT,
        rule: 8,
        explanation: "A chart conclusion and emphasized series require content judgment. Lucid Slides does not alter chart data or styling in local mode.",
        actionable: false,
        source: "local-analysis",
      });
    }

    const bulletCount = slide.elements.reduce(
      (sum, element) => sum + element.paragraphs.filter((paragraph) => paragraph.isBullet).length,
      0,
    );
    if (bulletCount >= 3) {
      findings.push({
        id: `slide-${slide.slide}-animation-rules-7-11`,
        slide: slide.slide,
        objectId: null,
        elementName: "Slide animation sequence",
        elementType: "animation",
        originalText: `${bulletCount} bullet paragraphs detected; all animation data is preserved.`,
        proposedText: null,
        placeholderText: "Manual PowerPoint review required — browser-safe animation authoring is not supported.",
        rule: 7,
        explanation: "Progressive reveal and appear/fade animation authoring are not implemented because reliable browser-side PowerPoint animation mutation is unavailable.",
        actionable: false,
        source: "local-analysis",
      });
    }
  }
  return findings;
}

async function sha256(bytes) {
  const digest = await globalThis.crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function assertSafePackageNames(names) {
  for (const name of names) {
    if (name.startsWith("/") || name.split("/").includes("..")) {
      throw new Error("The presentation contains an unsafe package path.");
    }
  }
}

function countMatching(names, pattern) {
  return names.filter((name) => pattern.test(name)).length;
}

export async function analyzePptx(arrayBuffer, onProgress, zipLibrary = globalThis.JSZip) {
  if (!zipLibrary?.loadAsync) throw new Error("The safe PowerPoint reader could not be loaded.");
  const sourceBytes = arrayBuffer instanceof Uint8Array
    ? arrayBuffer.slice()
    : new Uint8Array(arrayBuffer.slice(0));
  const sourceHash = await sha256(sourceBytes);
  onProgress?.("Validating the presentation package…");
  const zip = await zipLibrary.loadAsync(sourceBytes);
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  assertSafePackageNames(names);
  for (const part of REQUIRED_PARTS) {
    if (!zip.file(part)) throw new Error(`Missing required PowerPoint package part: ${part}`);
  }

  const slideFiles = names
    .map((name) => ({ name, match: name.match(SLIDE_PATH) }))
    .filter((entry) => entry.match)
    .sort((a, b) => Number(a.match[1]) - Number(b.match[1]));
  if (!slideFiles.length) throw new Error("No slides were found in this PowerPoint file.");

  const slides = [];
  const rawSlideHashes = {};
  for (const entry of slideFiles) {
    const slideNumber = Number(entry.match[1]);
    onProgress?.(`Analyzing slide ${slideNumber} without modifying it…`);
    const rawBytes = await zip.file(entry.name).async("uint8array");
    const xml = new TextDecoder().decode(rawBytes);
    rawSlideHashes[entry.name] = await sha256(rawBytes);
    slides.push(analyzeSlideXml(xml, slideNumber));
  }

  onProgress?.("Confirming that the source bytes are unchanged…");
  const finalHash = await sha256(sourceBytes);
  if (finalHash !== sourceHash) throw new Error("Safety validation failed: source bytes changed during analysis.");

  const inventory = {
    packageEntries: names.length,
    slides: slideFiles.length,
    media: countMatching(names, /^ppt\/media\//),
    slideRelationships: countMatching(names, /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/),
    notes: countMatching(names, /^ppt\/notesSlides\/notesSlide\d+\.xml$/),
    charts: countMatching(names, /^ppt\/charts\/chart\d+\.xml$/),
    tables: slides.reduce((sum, slide) => sum + slide.counts.tables, 0),
    hyperlinks: slides.reduce((sum, slide) => sum + slide.counts.hyperlinks, 0),
    imagesReferenced: slides.reduce((sum, slide) => sum + slide.counts.images, 0),
    words: slides.reduce((sum, slide) => sum + slide.counts.words, 0),
  };

  return {
    mode: "analysis-only",
    sourceHash,
    sourceBytes: sourceBytes.byteLength,
    packageValid: true,
    sourceUnchanged: true,
    outputPptxCreated: false,
    rawSlideHashes,
    inventory,
    slides,
    findings: buildLocalFindings(slides),
    limitations: [
      "No PowerPoint content, formatting, relationships, media, notes, charts, tables, hyperlinks, or structure was changed.",
      "No modified PowerPoint file was generated because reliable browser-side mutation has not been proven safe.",
      "Rules 7 and 11 require manual animation authoring in PowerPoint-compatible software.",
    ],
  };
}

export function createAnalysisSnapshot(analysis) {
  return {
    sourceHash: analysis.sourceHash,
    slides: analysis.slides.slice(0, 80).map((slide) => ({
      slide: slide.slide,
      elements: slide.elements.slice(0, 120).map((element) => ({
        objectId: element.objectId,
        name: element.name.slice(0, 160),
        type: element.type,
        text: element.text.slice(0, 6000),
      })),
    })),
  };
}

export function validateAiProposals(snapshot, proposals) {
  const lookup = new Map();
  for (const slide of snapshot.slides || []) {
    for (const element of slide.elements || []) lookup.set(`${slide.slide}:${element.objectId}`, element);
  }
  if (!Array.isArray(proposals)) return [];
  const validated = [];
  for (const proposal of proposals.slice(0, 200)) {
    const slide = Number(proposal.slide);
    const objectId = String(proposal.objectId || "");
    const element = lookup.get(`${slide}:${objectId}`);
    const originalText = String(proposal.originalText || "");
    const proposedText = String(proposal.proposedText || "").trim();
    const explanation = String(proposal.explanation || "").trim();
    const rule = Number(proposal.rule);
    if (!element || !originalText || !element.text.includes(originalText)) continue;
    if (!proposedText || proposedText === originalText || proposedText.length > 1200) continue;
    if (!explanation || explanation.length > 1000 || !Number.isInteger(rule) || rule < 1 || rule > 12) continue;
    validated.push({
      id: `ai-slide-${slide}-element-${objectId}-rule-${rule}-${validated.length + 1}`,
      slide,
      objectId,
      elementName: element.name,
      elementType: element.type,
      originalText,
      proposedText,
      rule,
      explanation,
      actionable: true,
      source: "openai-analysis",
      decision: "pending",
    });
  }
  return validated;
}

export async function simplifyPptx() {
  throw new Error(
    "PowerPoint mutation is disabled. Lucid Slides now operates in analysis-only mode and does not create a modified PPTX.",
  );
}

export { PLACEHOLDER_TEXT };
