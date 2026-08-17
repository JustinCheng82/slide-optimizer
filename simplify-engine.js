/* Lucid Slides - client-side PPTX simplification engine.
   Runs entirely in the browser (no server, no upload of your file anywhere).
   Requires JSZip to already be loaded as a global (window.JSZip).
*/

const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

const RULE_CONFIG = {
  maxWordsPerLine: 12,
  maxBoldWords: 4,
  boldFraction: 0.2,
  minLinesForFlag: 3,
  tinyImageEmuThreshold: 457200, // ~0.5 inch, EMUs (914400 EMU = 1 inch)
  maxStatFontHundredths: 9600, // cap enlarged stats at 96pt so they can't blow up
};

// Table cells hold factual/tabular data - trimming or rewording them risks
// breaking the table's meaning, so paragraphs inside <a:tc> are left alone.
function isInsideTableCell(node) {
  let cur = node.parentNode;
  while (cur && cur.nodeType === 1) {
    if (cur.tagName === "a:tc") return true;
    cur = cur.parentNode;
  }
  return false;
}

// Never rebuild a paragraph that contains a hyperlink run - collapsing runs
// would silently delete the link.
function hasHyperlinkRun(runs) {
  return runs.some((r) => {
    const rPr = r.getElementsByTagName("a:rPr")[0];
    return rPr && rPr.getElementsByTagName("a:hlinkClick").length > 0;
  });
}

async function simplifyPptx(arrayBuffer, onProgress) {
  const zip = await JSZip.loadAsync(arrayBuffer);

  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)\.xml/)[1], 10);
      const nb = parseInt(b.match(/slide(\d+)\.xml/)[1], 10);
      return na - nb;
    });

  if (!slideFiles.length) {
    throw new Error("Couldn't find any slides in this file - is it a valid .pptx?");
  }

  const report = [];

  for (const path of slideFiles) {
    const slideNum = parseInt(path.match(/slide(\d+)\.xml/)[1], 10);
    if (onProgress) onProgress(`Reading slide ${slideNum}...`);

    const xmlText = await zip.file(path).async("string");
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, "application/xml");

    if (doc.getElementsByTagName("parsererror").length) {
      report.push({ slide: slideNum, changes: [], flags: [`Could not parse this slide's XML - left untouched.`] });
      continue;
    }

    const slideReport = { slide: slideNum, changes: [], flags: [] };
    let totalNonEmptyLines = 0;

    const paragraphs = Array.from(doc.getElementsByTagName("a:p"));
    let hyperlinksPreserved = 0;
    for (const p of paragraphs) {
      const runs = Array.from(p.getElementsByTagName("a:r"));
      if (!runs.length) continue;
      if (isInsideTableCell(p)) continue;

      let fullText = "";
      runs.forEach((r) => {
        const tNodes = r.getElementsByTagName("a:t");
        if (tNodes.length) fullText += tNodes[0].textContent;
      });
      if (!fullText.trim()) continue;
      totalNonEmptyLines++;

      if (hasHyperlinkRun(runs)) {
        hyperlinksPreserved++;
        continue;
      }

      const trimmedText = fullText.trim();
      const words = trimmedText.split(/\s+/);
      const isStandaloneStat = words.length === 1 && /^\$?\d[\d,.]*%?$/.test(trimmedText);

      if (isStandaloneStat) {
        // Rule 6: make the key statistic bigger. Adding real context text
        // ("+34% engagement" instead of "34%") needs to know what the stat
        // means, so that part is left for a human - only sizing is automatic.
        runs.forEach((r) => {
          let rPr = r.getElementsByTagName("a:rPr")[0];
          if (!rPr) {
            rPr = doc.createElementNS(A_NS, "a:rPr");
            r.insertBefore(rPr, r.firstChild);
          }
          const currentSz = rPr.getAttribute("sz");
          const rawSz = currentSz ? Math.round(parseInt(currentSz, 10) * 1.5) : 4500;
          const newSz = Math.min(rawSz, RULE_CONFIG.maxStatFontHundredths);
          rPr.setAttribute("sz", String(newSz));
        });
        slideReport.changes.push(`Enlarged the standalone statistic "${trimmedText}".`);
        slideReport.flags.push(`Add context to "${trimmedText}" by hand, e.g. "+${trimmedText} engagement" (rule 6 needs to know what the number means).`);
        continue;
      }

      let newFullText = fullText;
      let trimmed = false;
      if (words.length > RULE_CONFIG.maxWordsPerLine) {
        newFullText = words.slice(0, RULE_CONFIG.maxWordsPerLine).join(" ") + "...";
        trimmed = true;
      }

      const newWords = newFullText.replace(/\.\.\.$/, "").trim().split(/\s+/);
      const boldCount =
        newWords.length >= 3
          ? Math.min(RULE_CONFIG.maxBoldWords, Math.ceil(newWords.length * RULE_CONFIG.boldFraction) + 1)
          : 0;

      if (trimmed || boldCount > 0) {
        const templateRPr = runs[0].getElementsByTagName("a:rPr")[0] || null;
        const endParaRPr = p.getElementsByTagName("a:endParaRPr")[0] || null;

        runs.forEach((r) => p.removeChild(r));

        const makeRun = (text, bold) => {
          const r = doc.createElementNS(A_NS, "a:r");
          if (templateRPr || bold) {
            const rPr = templateRPr ? templateRPr.cloneNode(true) : doc.createElementNS(A_NS, "a:rPr");
            if (bold) rPr.setAttribute("b", "1");
            r.appendChild(rPr);
          }
          const t = doc.createElementNS(A_NS, "a:t");
          t.textContent = text;
          r.appendChild(t);
          return r;
        };

        if (boldCount > 0) {
          const boldPhrase = newWords.slice(0, boldCount).join(" ");
          const rest = newFullText.slice(boldPhrase.length);
          p.insertBefore(makeRun(boldPhrase, true), endParaRPr);
          if (rest.length) p.insertBefore(makeRun(rest, false), endParaRPr);
        } else {
          p.insertBefore(makeRun(newFullText, false), endParaRPr);
        }

        if (trimmed) slideReport.changes.push(`Trimmed a line to ${RULE_CONFIG.maxWordsPerLine} words.`);
        if (boldCount > 0) slideReport.changes.push(`Bolded the first ${boldCount} word(s) of a line.`);
      }
    }

    if (hyperlinksPreserved) {
      slideReport.flags.push(`${hyperlinksPreserved} line(s) contain a link - left completely untouched so the link isn't broken. Trim/bold those by hand if needed.`);
    }

    // Rule 8: flag charts for manual attention (can't safely script chart-data edits).
    const chartFrames = Array.from(doc.getElementsByTagName("p:graphicFrame")).filter(
      (gf) => gf.getElementsByTagName("c:chart").length > 0
    );
    if (chartFrames.length) {
      slideReport.flags.push(`Has a chart - manually emphasize the data that supports your takeaway, fade the rest, and change the chart title into a conclusion (rule 8). Reading and rewriting chart data needs a human.`);
    }

    // Rule 9: drop small images that read as decorative.
    const pics = Array.from(doc.getElementsByTagName("p:pic"));
    let removedImages = 0;
    pics.forEach((pic) => {
      const ext = pic.getElementsByTagName("a:ext")[0];
      if (!ext) return;
      const cx = parseInt(ext.getAttribute("cx") || "0", 10);
      const cy = parseInt(ext.getAttribute("cy") || "0", 10);
      if (cx > 0 && cy > 0 && cx < RULE_CONFIG.tinyImageEmuThreshold && cy < RULE_CONFIG.tinyImageEmuThreshold) {
        pic.parentNode.removeChild(pic);
        removedImages++;
      }
    });
    if (removedImages) slideReport.changes.push(`Removed ${removedImages} small image(s) that looked decorative.`);

    if (totalNonEmptyLines >= RULE_CONFIG.minLinesForFlag) {
      slideReport.flags.push(
        `${totalNonEmptyLines} lines of text on one slide - add a progressive "Appear" build in PowerPoint's Animations pane (rules 7 & 11), or split into multiple slides (rule 10). Both need your judgment, so they're not done automatically.`
      );
    }
    slideReport.flags.push(`Confirm the title states the one main takeaway in <=10 words (rules 1-2), and that one element is clearly the most visually dominant thing on the slide (rule 5).`);

    if (slideReport.changes.length || slideReport.flags.length) report.push(slideReport);

    const serializer = new XMLSerializer();
    zip.file(path, serializer.serializeToString(doc));
  }

  if (onProgress) onProgress("Packaging the simplified file...");
  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });

  return { blob, report };
}
