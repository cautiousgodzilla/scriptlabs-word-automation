export const BUILTIN_WORKFLOWS = [
  {
    id: "builtin-summarize-selection",
    title: "Summarize selection",
    type: "assistant",
    practice: "general",
    prompt_md: `Summarize the selected Word text into 3–5 concise bullet points covering the key ideas. Be factual and direct. Put the summary in your message, then insert the bullets below the current selection in the Word document using Word.run().`,
  },
  {
    id: "builtin-rewrite-formal",
    title: "Rewrite (formal tone)",
    type: "assistant",
    practice: "writing",
    prompt_md: `Rewrite the selected text in a formal, professional tone. Preserve the meaning exactly — replace colloquial phrases with precise language. Return the rewritten version in your message, then replace the current Word selection with it using Word.run().`,
  },
  {
    id: "builtin-extract-action-items",
    title: "Extract action items",
    type: "assistant",
    practice: "productivity",
    prompt_md: `Read the selected text and extract every action item, task, or commitment mentioned. Present them as a numbered list in your message, then insert that list below the selection in the Word document using Word.run().`,
  },
  {
    id: "builtin-explain-plain-english",
    title: "Explain in plain English",
    type: "assistant",
    practice: "general",
    prompt_md: `Explain the selected text in plain, simple English for someone unfamiliar with the subject. Keep it to 2–3 short paragraphs. Reply in the chat only — do not modify the document.`,
  },
  {
    id: "builtin-flowchart-from-steps",
    title: "Generate flowchart",
    type: "assistant",
    practice: "visual",
    prompt_md: `## Generate Flowchart Workflow

The server will generate all Office.js code — your job is ONLY to extract structured data and return it.
Do NOT put any code in "actions". The server reads "flowchart_steps" and renders the flowchart itself
with guaranteed styling (white fill, black text, lead lines, reference numerals).

### Step 1 — Find the source text
Priority order:
1. \`context.selection\` — highlighted text in Word. Use as-is.
2. \`context.bodyText\` — if selection is empty, scan the document body for numbered steps, procedures, or any sequential process.
3. If the user's message names a topic, search bodyText for that topic.
4. If none yields processable content, set \`message\` to ask the user to select the relevant text and return no \`flowchart_steps\`.

### Step 2 — Extract and label steps
Parse the source into 3–10 ordered steps:
- Labels: ≤ 6 words each.
- Shape type for each step:
  - First node → \`"FlowChartTerminator"\` labelled "Start"
  - Last node  → \`"FlowChartTerminator"\` labelled "End"
  - Steps with if / when / check / decide / is / does language → \`"FlowChartDecision"\`
  - All other steps → \`"FlowChartProcess"\`
- Merge trivial sub-points into the nearest parent step.

### Step 3 — Determine reference numerals
Scan the user's message for an explicit numeral list (e.g. "use A01, A03, A05" or "reference numbers 101, 103, 105").
- If found: use those values in order, one per step.
- If not found: auto-generate X02, X04, X06, X08 … (increment by 2, prefix "X").
Produce exactly as many numerals as there are steps.

### Step 4 — Return ONLY structured data

Return this exact JSON shape — no "actions" field, no code:

{
  "message": "Identified N steps: Start (X02), … End (XNN). Flowchart will be drawn on the active slide.",
  "flowchart_steps": [
    { "label": "Start",        "type": "FlowChartTerminator", "ref": "X02" },
    { "label": "Your step",    "type": "FlowChartProcess",    "ref": "X04" },
    { "label": "Decision?",    "type": "FlowChartDecision",   "ref": "X06" },
    { "label": "End",          "type": "FlowChartTerminator", "ref": "X08" }
  ]
}

Use the actual extracted step labels, types, and reference numerals. Do not include "actions".`,
  },
  {
    id: "builtin-key-questions",
    title: "Generate key questions",
    type: "assistant",
    practice: "research",
    prompt_md: `Read the selected text and generate 5–8 incisive questions a critical reader would ask about it. Format them as a numbered list in your message. Do not modify the document.`,
  },
  {
    id: "builtin-proofread-document",
    title: "Proofread document",
    type: "assistant",
    practice: "writing",
    prompt_md: `## Proofread Document Workflow

Use \`context.fullBody\` as the document text. Do NOT ask the user to select text or confirm anything — always use the full body automatically. Fall back to \`context.bodyText\` then \`context.selection\` only if \`fullBody\` is absent.

### What to check:
1. Grammar & spelling — wrong words, subject-verb disagreement, spelling errors, punctuation.
2. Clarity — ambiguous sentences, unclear antecedents.
3. Consistency — capitalisation, hyphenation, terminology, number formatting.
4. Tone & style — passive voice overuse, wordy phrases ("due to the fact that" → "because").
5. Structure — missing transitions, orphaned points.

### YOU MUST return this exact JSON — no plain text, no refusals, no "actions" field:

{
  "message": "Found N issues. Applying tracked changes with comments.",
  "proofreading_changes": [
    { "find": "exact verbatim phrase from document", "replace": "corrected version", "reason": "Grammar: subject-verb agreement" }
  ]
}

If the document is empty or has no errors, return:
{ "message": "Document is clean — no issues found.", "proofreading_changes": [] }

### Rules for each entry:
- **find**: copy the exact phrase from the document (same capitalisation, same punctuation). Keep to 5–20 words.
- **replace**: the corrected version of that phrase. Must differ from find.
- **reason**: one short label, e.g. "Spelling: 'recieve'", "Style: wordy".
- Max 10 entries. No duplicate find values. No entry where find === replace. No "actions" field.`,
  },
];
