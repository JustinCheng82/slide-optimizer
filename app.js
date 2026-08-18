import {
  analyzePptx,
  createAnalysisSnapshot,
  validateAiProposals,
} from "./simplify-engine.js";

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const processingBox = document.getElementById("processingBox");
const errorBox = document.getElementById("errorBox");
const resultBox = document.getElementById("resultBox2");
const resultTitle = document.getElementById("resultTitle");
const resultSummary = document.getElementById("resultSummary");
const modeNotice = document.getElementById("modeNotice");
const inventoryList = document.getElementById("inventoryList");
const proposalContainer = document.getElementById("proposalContainer");
const reportDownload = document.getElementById("reportDownload");

let currentReportUrl = "";
let currentResult = null;

function resetPanels() {
  processingBox.style.display = "none";
  processingBox.replaceChildren();
  errorBox.style.display = "none";
  errorBox.textContent = "";
  resultBox.style.display = "none";
  proposalContainer.replaceChildren();
  inventoryList.replaceChildren();
  if (currentReportUrl) URL.revokeObjectURL(currentReportUrl);
  currentReportUrl = "";
  currentResult = null;
}

function logLine(message) {
  processingBox.style.display = "block";
  const line = document.createElement("div");
  line.className = "log-line";
  line.textContent = message;
  processingBox.appendChild(line);
}

function showError(message) {
  errorBox.style.display = "block";
  errorBox.textContent = message;
}

function appendInventory(label, value) {
  const item = document.createElement("li");
  const strong = document.createElement("strong");
  strong.textContent = `${value} `;
  item.append(strong, document.createTextNode(label));
  inventoryList.appendChild(item);
}

function updateReportDownload() {
  if (!currentResult) return;
  const report = {
    generatedAt: new Date().toISOString(),
    mode: "analysis-only",
    fileName: currentResult.fileName,
    sourceHash: currentResult.analysis.sourceHash,
    sourceUnchanged: currentResult.analysis.sourceUnchanged,
    outputPptxCreated: false,
    inventory: currentResult.analysis.inventory,
    proposals: currentResult.items.map((item) => ({
      slide: item.slide,
      objectId: item.objectId,
      elementName: item.elementName,
      originalText: item.originalText,
      proposedText: item.proposedText || item.placeholderText,
      rule: item.rule,
      explanation: item.explanation,
      source: item.source,
      decision: item.decision || "not-actionable",
    })),
    limitations: currentResult.analysis.limitations,
  };
  if (currentReportUrl) URL.revokeObjectURL(currentReportUrl);
  currentReportUrl = URL.createObjectURL(
    new Blob([JSON.stringify(report, null, 2)], { type: "application/json" }),
  );
  reportDownload.href = currentReportUrl;
  reportDownload.download = currentResult.fileName.replace(/\.pptx$/i, "") + " — Lucid Slides analysis.json";
}

function makeTextBlock(label, text, className) {
  const block = document.createElement("div");
  block.className = className;
  const heading = document.createElement("strong");
  heading.textContent = label;
  const value = document.createElement("p");
  value.textContent = text || "None";
  block.append(heading, value);
  return block;
}

function renderItems(items) {
  proposalContainer.replaceChildren();
  if (!items.length) {
    const empty = document.createElement("p");
    empty.className = "empty-findings";
    empty.textContent = "No review candidates were found. Your presentation was still left completely unchanged.";
    proposalContainer.appendChild(empty);
    return;
  }

  for (const item of items) {
    const card = document.createElement("article");
    card.className = "proposal-card";
    card.dataset.proposalId = item.id;

    const heading = document.createElement("h4");
    heading.textContent = `Slide ${item.slide} · ${item.elementName || `Element ${item.objectId || "unknown"}`}`;
    const meta = document.createElement("p");
    meta.className = "proposal-meta";
    meta.textContent = `Rule ${item.rule} · ${item.source === "openai-analysis" ? "AI proposal" : "local review finding"}`;

    card.append(
      heading,
      meta,
      makeTextBlock("Original", item.originalText, "proposal-text original-text"),
      makeTextBlock(
        "Proposed",
        item.proposedText || item.placeholderText,
        `proposal-text proposed-text${item.actionable ? "" : " placeholder-proposal"}`,
      ),
      makeTextBlock("Why this rule applies", item.explanation, "proposal-explanation"),
    );

    const controls = document.createElement("div");
    controls.className = "proposal-controls";
    if (item.actionable) {
      for (const decision of ["approved", "rejected"]) {
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = decision === "approved" ? "Approve proposal" : "Reject proposal";
        button.className = decision === "approved" ? "approve-btn" : "reject-btn";
        button.addEventListener("click", () => {
          item.decision = decision;
          card.dataset.decision = decision;
          controls.querySelectorAll("button").forEach((control) => control.removeAttribute("aria-pressed"));
          button.setAttribute("aria-pressed", "true");
          updateReportDownload();
        });
        controls.appendChild(button);
      }
    } else {
      const note = document.createElement("span");
      note.className = "not-actionable";
      note.textContent = "No edit proposed — approval is disabled until a meaning-based proposal exists.";
      controls.appendChild(note);
    }
    card.appendChild(controls);
    proposalContainer.appendChild(card);
  }
}

async function requestAiProposals(analysis) {
  const snapshot = createAnalysisSnapshot(analysis);
  try {
    const response = await fetch("/api/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Lucid-Request": "analysis-v1" },
      body: JSON.stringify({ presentation: snapshot }),
    });
    const data = await response.json().catch(() => ({}));
    if (response.status === 503 && data.mode === "analysis-only") {
      return { proposals: [], notice: data.message || "AI analysis is not configured." };
    }
    if (!response.ok) throw new Error(data.error || `Backend request failed (${response.status})`);
    return {
      proposals: validateAiProposals(snapshot, data.proposals),
      notice: "AI-generated proposals were validated against exact slide and element IDs. Nothing was applied.",
    };
  } catch (error) {
    return {
      proposals: [],
      notice: `AI analysis was unavailable (${error.message}). Showing local findings only; no text was generated or changed.`,
    };
  }
}

async function handleFile(file) {
  resetPanels();
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".pptx")) {
    showError("Choose a .pptx file. Lucid Slides reads it locally for analysis and does not create a modified presentation.");
    return;
  }

  try {
    logLine(`Reading ${file.name} locally…`);
    const arrayBuffer = await file.arrayBuffer();
    const analysis = await analyzePptx(arrayBuffer, logLine);
    logLine("Requesting optional meaning-based proposals…");
    const ai = await requestAiProposals(analysis);
    const items = [...ai.proposals, ...analysis.findings];
    currentResult = { fileName: file.name, analysis, items };

    resultTitle.textContent = "Analysis complete — your presentation was not modified.";
    resultSummary.textContent = `${analysis.inventory.slides} slides and ${analysis.inventory.words.toLocaleString()} words were inspected. No PowerPoint output was created.`;
    modeNotice.textContent = ai.notice;
    appendInventory("slides preserved", analysis.inventory.slides);
    appendInventory("media files preserved", analysis.inventory.media);
    appendInventory("slide relationship parts preserved", analysis.inventory.slideRelationships);
    appendInventory("notes parts preserved", analysis.inventory.notes);
    appendInventory("charts preserved", analysis.inventory.charts);
    appendInventory("tables detected and preserved", analysis.inventory.tables);
    appendInventory("hyperlinked paragraphs detected and preserved", analysis.inventory.hyperlinks);
    renderItems(items);
    updateReportDownload();

    processingBox.style.display = "none";
    resultBox.style.display = "block";
    resultBox.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    console.error(error);
    processingBox.style.display = "none";
    showError(`Lucid Slides could not safely analyze that file: ${error.message}`);
  }
}

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});
fileInput.addEventListener("change", (event) => handleFile(event.target.files[0]));

for (const eventName of ["dragenter", "dragover"]) {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
}
for (const eventName of ["dragleave", "drop"]) {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
  });
}
dropzone.addEventListener("drop", (event) => handleFile(event.dataTransfer.files[0]));
