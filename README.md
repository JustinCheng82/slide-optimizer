# Lucid Slides

Lucid Slides is a browser-based, read-only PowerPoint analyzer for presentation clarity. The production safety mode does **not** modify or regenerate `.pptx` files.

## Safety status

The previous engine was unsafe. It truncated paragraphs, bolded opening words, rebuilt run-level formatting, removed images by dimensions, serialized every slide XML file, and treated ZIP creation as success. Those behaviors are disabled.

The repaired workflow:

1. Reads a selected `.pptx` locally with JSZip.
2. Validates required package parts and unsafe paths.
3. Extracts slide and element identifiers, exact text, and preservation inventories without serializing XML.
4. Shows local review findings with explicit “AI analysis required” placeholders.
5. Optionally sends only a reduced text-and-ID snapshot—not the PowerPoint file—to `/api/analyze`.
6. Shows validated proposals for individual approval or rejection.
7. Downloads a JSON analysis report only. It never downloads a modified presentation.

No PowerPoint content, images, relationships, notes, charts, tables, hyperlinks, text runs, fonts, formatting, or slide structure is written by the application.

## Optional OpenAI analysis

`api/analyze.js` is a Vercel Function that keeps `OPENAI_API_KEY` server-side, sends requests with `store: false`, requires structured output, and revalidates every proposal against exact slide and object IDs and source text. Google or PowerPoint credentials are not required for local analysis.

If `OPENAI_API_KEY` is absent, the endpoint returns a clearly labeled `analysis-only` response and the browser continues with local findings. Never put an API key in frontend code or ask a user to paste one into the site.

Environment variables are documented in `.env.example`.

## Run locally

```bash
python3 -m http.server 8000
```

Open `http://localhost:8000`. The static server is only for local development; users access the deployed website normally.

## Tests

```bash
npm test
```

Run the full supplied-deck regression on this Mac with:

```bash
LUCID_TEST_PPTX="/Users/nathan/Desktop/slide for test.pptx" npm test
```

The real-deck test confirms 22 slides are inspected, source bytes and every raw slide hash remain unchanged, and no PowerPoint output is created. Because analysis-only mode produces no `.pptx`, output-package and PowerPoint repair-warning checks are intentionally not applicable; the validated input remains the only presentation file.

## Known platform limitation

Reliable in-browser mutation of arbitrary PowerPoint packages has not been established. In particular, broad DOM parsing and serialization can damage compatibility even when media and relationships remain in the ZIP. Targeted application of approved edits must remain disabled until a future implementation can preserve untouched ZIP entries byte-for-byte and pass visual and PowerPoint-compatible open/repair testing on representative decks.
