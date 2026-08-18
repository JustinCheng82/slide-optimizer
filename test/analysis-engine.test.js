import assert from "node:assert/strict";
import { createHash, webcrypto } from "node:crypto";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  analyzePptx,
  analyzeSlideXml,
  buildLocalFindings,
  simplifyPptx,
  validateAiProposals,
} from "../simplify-engine.js";
import { makeFixturePptx, loadJsZip } from "./helpers.js";

if (!globalThis.crypto) globalThis.crypto = webcrypto;

function digest(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

test("slide inspection preserves exact text and identifies existing elements", () => {
  const xml = `<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="7" name="Body"/></p:nvSpPr><p:txBody><a:p><a:r><a:rPr b="1"/><a:t>First </a:t></a:r><a:r><a:rPr baseline="30000"/><a:t>formatted</a:t></a:r><a:r><a:t> words remain.</a:t></a:r><a:hlinkClick r:id="rId8"/></a:p></p:txBody></p:sp>
    <p:pic><p:nvPicPr><p:cNvPr id="8" name="Small but meaningful logo"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic>
  </p:spTree></p:cSld></p:sld>`;
  const slide = analyzeSlideXml(xml, 1);
  assert.equal(slide.elements[0].text, "First formatted words remain.");
  assert.equal(slide.elements[0].paragraphs[0].hasHyperlink, true);
  assert.equal(slide.elements[1].name, "Small but meaningful logo");
  assert.equal(slide.elements[1].relationshipId, "rId3");
  assert.equal(slide.counts.images, 1);
});

test("local findings never invent text or propose image deletion", () => {
  const findings = buildLocalFindings([{
    slide: 1,
    elements: [{
      objectId: "2",
      name: "Body",
      type: "text",
      text: "This paragraph contains more than twelve words but every original word must remain completely intact.",
      paragraphs: [{ index: 0, text: "This paragraph contains more than twelve words but every original word must remain completely intact.", wordCount: 14, boldWordCount: 0, hasHyperlink: false, isBullet: false }],
    }, { objectId: "3", name: "Logo", type: "image", text: "", paragraphs: [] }],
    counts: { charts: 0 },
  }]);
  assert.equal(findings.length, 1);
  assert.equal(findings[0].proposedText, null);
  assert.equal(findings[0].actionable, false);
  assert.equal(findings.some((finding) => /delete|remove/i.test(finding.explanation)), false);
});

test("analysis preserves text, media, relationships, hyperlinks, notes, charts, tables, and source bytes", async () => {
  const { JSZip, bytes } = await makeFixturePptx();
  const beforeHash = digest(bytes);
  const beforeZip = await JSZip.loadAsync(bytes);
  const protectedParts = [
    "ppt/slides/slide1.xml",
    "ppt/slides/_rels/slide1.xml.rels",
    "ppt/media/image1.png",
    "ppt/notesSlides/notesSlide1.xml",
    "ppt/charts/chart1.xml",
  ];
  const beforeParts = new Map(await Promise.all(protectedParts.map(async (name) => [name, digest(await beforeZip.file(name).async("uint8array"))])));

  const analysis = await analyzePptx(bytes, null, JSZip);
  assert.equal(digest(bytes), beforeHash);
  assert.equal(analysis.packageValid, true);
  assert.equal(analysis.sourceUnchanged, true);
  assert.equal(analysis.outputPptxCreated, false);
  assert.equal(analysis.inventory.slides, 1);
  assert.equal(analysis.inventory.media, 1);
  assert.equal(analysis.inventory.slideRelationships, 1);
  assert.equal(analysis.inventory.notes, 1);
  assert.equal(analysis.inventory.charts, 1);
  assert.equal(analysis.inventory.tables, 1);
  assert.equal(analysis.inventory.hyperlinks, 1);

  const afterZip = await JSZip.loadAsync(bytes);
  for (const [name, hash] of beforeParts) {
    assert.equal(digest(await afterZip.file(name).async("uint8array")), hash, `${name} changed`);
  }
});

test("unapproved PowerPoint mutation is hard-disabled", async () => {
  await assert.rejects(() => simplifyPptx(new ArrayBuffer(0)), /mutation is disabled/i);
});

test("production engine contains no broad XML serialization or PPTX generation path", async () => {
  const source = await readFile(new URL("../simplify-engine.js", import.meta.url), "utf8");
  assert.doesNotMatch(source, /XMLSerializer|generateAsync|removeChild|tinyImageEmuThreshold|maxWordsPerLine|maxBoldWords/);
  assert.match(source, /outputPptxCreated:\s*false/);
});

test("AI proposals must reference exact existing slide text", () => {
  const snapshot = { slides: [{ slide: 2, elements: [{ objectId: "9", name: "Body", type: "text", text: "Exact original sentence." }] }] };
  const proposals = validateAiProposals(snapshot, [
    { slide: 2, objectId: "9", originalText: "Exact original sentence.", proposedText: "Clearer sentence.", rule: 3, explanation: "Removes repetition after semantic review." },
    { slide: 2, objectId: "missing", originalText: "Invented", proposedText: "Bad", rule: 3, explanation: "Invalid." },
  ]);
  assert.equal(proposals.length, 1);
  assert.equal(proposals[0].decision, "pending");
});

test("supplied 22-slide deck remains byte-identical during complete analysis", { skip: !process.env.LUCID_TEST_PPTX }, async () => {
  const bytes = new Uint8Array(await readFile(process.env.LUCID_TEST_PPTX));
  const JSZip = await loadJsZip();
  const beforeHash = digest(bytes);
  const analysis = await analyzePptx(bytes, null, JSZip);
  assert.equal(beforeHash, "967649a18c96fd0fb8b522abc0a4c4dbb50e5ca493728c512d5f49cf1ae1a68c");
  assert.equal(analysis.inventory.slides, 22);
  assert.equal(analysis.inventory.media, 26);
  assert.equal(analysis.inventory.slideRelationships, 22);
  assert.equal(analysis.inventory.notes, 22);
  assert.equal(analysis.inventory.imagesReferenced, 15);
  assert.equal(analysis.inventory.words, 1577);
  assert.equal(digest(bytes), beforeHash);
  assert.equal(analysis.sourceUnchanged, true);
  assert.equal(analysis.outputPptxCreated, false);
  assert.equal(Object.keys(analysis.rawSlideHashes).length, 22);
});
