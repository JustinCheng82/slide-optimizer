import assert from "node:assert/strict";
import test from "node:test";
import handler from "../api/analyze.js";
import { validateProposalResponse, validateSnapshot } from "../lib/proposals.js";

function mockResponse() {
  return {
    headers: {},
    statusCode: 200,
    setHeader(name, value) { this.headers[name] = value; },
    status(code) { this.statusCode = code; return this; },
    json(payload) { this.payload = payload; return this; },
  };
}

const snapshot = {
  sourceHash: "abc",
  slides: [{ slide: 1, elements: [{ objectId: "2", name: "Body", type: "text", text: "Original exact text." }] }],
};

test("missing OpenAI credentials returns explicit analysis-only mode", async () => {
  const previousKey = process.env.OPENAI_API_KEY;
  delete process.env.OPENAI_API_KEY;
  const request = {
    method: "POST",
    headers: { "x-lucid-request": "analysis-v1" },
    body: { presentation: snapshot },
    socket: { remoteAddress: "test-missing-key" },
  };
  const response = mockResponse();
  await handler(request, response);
  assert.equal(response.statusCode, 503);
  assert.equal(response.payload.mode, "analysis-only");
  assert.deepEqual(response.payload.proposals, []);
  if (previousKey) process.env.OPENAI_API_KEY = previousKey;
});

test("snapshot validation strips unknown fields and caps content", () => {
  const validated = validateSnapshot({ ...snapshot, secret: "discard", slides: [{ ...snapshot.slides[0], extra: "discard" }] });
  assert.deepEqual(Object.keys(validated).sort(), ["slides", "sourceHash"]);
  assert.deepEqual(Object.keys(validated.slides[0]).sort(), ["elements", "slide"]);
});

test("server rejects model proposals for invented elements or text", () => {
  const validatedSnapshot = validateSnapshot(snapshot);
  const proposals = validateProposalResponse(validatedSnapshot, { proposals: [
    { slide: 1, objectId: "2", originalText: "Original exact text.", proposedText: "A clearer version.", rule: 3, explanation: "Meaning-based rewrite." },
    { slide: 1, objectId: "2", originalText: "Invented source", proposedText: "Invalid", rule: 3, explanation: "Invalid." },
    { slide: 1, objectId: "99", originalText: "Original exact text.", proposedText: "Invalid", rule: 3, explanation: "Invalid." },
  ] });
  assert.equal(proposals.length, 1);
});
