/**
 * Model-agnostic LLM client with tool-calling support.
 * Reads LLM_PROVIDER, LLM_API_KEY, LLM_MODEL from environment.
 * Supports: anthropic, openai, google
 */

const PROVIDER = (process.env.LLM_PROVIDER || "").toLowerCase();
const API_KEY = process.env.LLM_API_KEY;
const MODEL = process.env.LLM_MODEL;

if (!PROVIDER) throw new Error("LLM_PROVIDER is not set");
if (!API_KEY) throw new Error("LLM_API_KEY is not set");
if (!MODEL) throw new Error("LLM_MODEL is not set");

const SYSTEM_PROMPT = `You are an Office automation agent. You receive messages from Microsoft Word and PowerPoint task panes running Script Lab.

EVERY reply MUST be valid JSON — no exceptions, no plain text, no apologies, no refusals.
The default schema is:
{
  "message": "<explanation or response to the user>",
  "actions": [
    {
      "target": "word" | "powerpoint" | "all",
      "code": "<Office.js code string, or null if no action>"
    }
  ]
}

IMPORTANT EXCEPTION: When a user message starts with "## Active Workflow:", read the workflow
instructions carefully. If they tell you to return a custom top-level field (e.g. "flowchart_steps"
or "proofreading_changes"), you MUST return exactly that field and MUST NOT add an "actions" field.
The server reads the custom field and generates all Office.js code itself.

If a workflow is active and the document content is empty or insufficient, return the custom field
with an empty array and explain in "message" — never refuse or reply with plain text.

Rules for the "code" field (when "actions" is used):
- Do not wrap code in markdown fences
- If no code action is needed, omit "actions" or set it to []
- The code runs inside Script Lab via new Function(), so Word, PowerPoint, Office globals are available on window
- Always use async/await; always call await ctx.sync() after every .load() before accessing properties
- For cross-app tasks, include one action per target app

## Step 1 — Understand intent before writing code

When the user refers to selected text or document content, first decide WHAT to do with it:

| Intent signal | What to generate |
|---|---|
| "copy", "paste", "add this", "put this" | Insert text verbatim as a text box |
| "summarize", "key points", "bullets" | Summarize the text, insert as bullet list |
| "flowchart", "diagram", "steps", "process" | Parse steps from text, build a flowchart with shapes + connectors |
| "table", "compare" | Build a PowerPoint table from the content |
| "title / heading" | Insert into the title placeholder of the slide |

## Step 2 — Use actual context values in code

The user message includes a Context JSON with "selection" and "bodyText".
NEVER use placeholder strings like "[Content from Word Document]". Always embed the real text.

## Word patterns

Word.run(async (ctx) => {
  const body = ctx.document.body;
  body.load("text");
  await ctx.sync();
  console.log(body.text);
});

Word.run(async (ctx) => {
  ctx.document.body.insertText("Hello", Word.InsertLocation.end);
  await ctx.sync();
});

## PowerPoint patterns

CRITICAL: Every collection must be .load()-ed then ctx.sync()-ed before iterating or reading properties.

// Read slide count
await PowerPoint.run(async (ctx) => {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  console.log("slides:", slides.items.length);
});

// Add a text box to slide 0
await PowerPoint.run(async (ctx) => {
  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();
  const slide = slides.items[0];
  slide.shapes.addTextBox("Hello World", { left: 100, top: 100, width: 300, height: 60 });
  await ctx.sync();
});

// Add a geometric shape — PascalCase strings, dimensions in the options object
await PowerPoint.run(async (ctx) => {
  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();
  // Valid types: "Rectangle","Ellipse","Diamond","RoundRectangle","FlowChartProcess","FlowChartDecision","FlowChartTerminator"
  const shape = slides.items[0].shapes.addGeometricShape("Rectangle", { left: 100, top: 100, width: 200, height: 60 });
  shape.textFrame.textRange.text = "Label";
  await ctx.sync();
});

// Add a new slide
await PowerPoint.run(async (ctx) => {
  ctx.presentation.slides.add();
  await ctx.sync();
});

## PowerPoint flowchart pattern

CRITICAL: Never use enum references (PowerPoint.GeometricShapeType.*, PowerPoint.ConnectorType.*) — they are undefined at runtime.
All shape type strings are PascalCase. "rectangle" is WRONG. "Rectangle" is correct.
Pass { left, top, width, height } as the second argument to addGeometricShape.
addLine in PowerPoint takes (connectorType, { left, top, width, height }) — bounding box, NOT beginX/beginY/endX/endY.

await PowerPoint.run(async (ctx) => {
  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();
  const slide = slides.items[0];

  const W = 220, H = 60, LEFT = 260;
  const steps = [
    { label: "Start",  type: "FlowChartTerminator", top: 40  },
    { label: "Step 1", type: "FlowChartProcess",    top: 130 },
    { label: "Step 2", type: "FlowChartProcess",    top: 220 },
    { label: "End",    type: "FlowChartTerminator", top: 310 },
  ];

  for (const s of steps) {
    const shape = slide.shapes.addGeometricShape(s.type, { left: LEFT, top: s.top, width: W, height: H });
    shape.textFrame.textRange.text = s.label;
  }

  for (let i = 0; i < steps.length - 1; i++) {
    const fromBottom = steps[i].top + H;
    const toTop = steps[i + 1].top;
    slide.shapes.addLine("Straight", { left: LEFT + W / 2 - 1, top: fromBottom, width: 2, height: toTop - fromBottom });
  }
  await ctx.sync();
});

## Workflows

When a user message contains a "## Active Workflow:" section, that section contains the complete
instructions for this turn. You MUST follow those instructions exactly — they override the default
JSON schema and all other defaults. In particular:
- If the instructions say to return a custom field (e.g. "flowchart_steps", "proofreading_changes"),
  return ONLY that field — do NOT add an "actions" field or generate any code yourself.
- If the instructions say the server generates the code, trust that and return only the structured data.
- Never refuse, never reply with plain text, and never ask the user for clarification when a workflow is active.
  If data is missing or empty, return the custom field with an empty array and explain in "message".

You may call list_workflows if the user explicitly asks what workflow templates exist.
Do NOT call read_workflow — the workflow content is already embedded in the message.`;

// ---------------------------------------------------------------------------
// Canonical tool definitions — converted per-provider below
// ---------------------------------------------------------------------------
const TOOL_DEFS = [
  {
    name: "list_workflows",
    description: "List all workflows available to the user. Call this when the user explicitly asks what templates or workflows exist.",
    parameters: { type: "object", properties: {} },
  },
];

const ANTHROPIC_TOOLS = TOOL_DEFS.map((t) => ({
  name: t.name,
  description: t.description,
  input_schema: t.parameters,
}));

const OPENAI_TOOLS = TOOL_DEFS.map((t) => ({
  type: "function",
  function: { name: t.name, description: t.description, parameters: t.parameters },
}));

const GOOGLE_TOOLS = [
  { functionDeclarations: TOOL_DEFS.map((t) => ({ name: t.name, description: t.description, parameters: t.parameters })) },
];

// ---------------------------------------------------------------------------
// Tool executor factory
// ---------------------------------------------------------------------------
function makeExecuteTool(workflowStore) {
  return async function executeTool(name) {
    if (name === "list_workflows") {
      return JSON.stringify(
        Array.from(workflowStore.entries()).map(([id, w]) => ({ id, title: w.title }))
      );
    }
    return "Unknown tool.";
  };
}

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------
export async function chat(conversationMessages, { workflowStore = new Map() } = {}) {
  const executeTool = makeExecuteTool(workflowStore);
  if (PROVIDER === "anthropic") return runAnthropic(conversationMessages, executeTool);
  if (PROVIDER === "openai") return runOpenAI(conversationMessages, executeTool);
  if (PROVIDER === "google") return runGoogle(conversationMessages, executeTool);
  throw new Error(`Unsupported LLM_PROVIDER: "${PROVIDER}". Use anthropic, openai, or google.`);
}

// ---------------------------------------------------------------------------
// Anthropic agentic loop
// ---------------------------------------------------------------------------
async function runAnthropic(messages, executeTool) {
  let loop = messages.map((m) => ({
    role: m.role,
    content: typeof m.content === "string" ? m.content : JSON.stringify(m.content),
  }));

  for (;;) {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 4096,
        system: SYSTEM_PROMPT,
        messages: loop,
        tools: ANTHROPIC_TOOLS,
      }),
    });
    if (!res.ok) throw new Error(`Anthropic error ${res.status}: ${await res.text()}`);
    const data = await res.json();

    if (data.stop_reason !== "tool_use") {
      return data.content.find((b) => b.type === "text")?.text ?? "";
    }

    const toolUseBlocks = data.content.filter((b) => b.type === "tool_use");
    const toolResults = [];
    for (const block of toolUseBlocks) {
      const result = await executeTool(block.name, block.input ?? {});
      toolResults.push({ type: "tool_result", tool_use_id: block.id, content: result });
    }

    loop = [...loop, { role: "assistant", content: data.content }, { role: "user", content: toolResults }];
  }
}

// ---------------------------------------------------------------------------
// OpenAI agentic loop
// ---------------------------------------------------------------------------
async function runOpenAI(messages, executeTool) {
  const apiMessages = [
    { role: "system", content: SYSTEM_PROMPT },
    ...messages.map((m) => ({
      role: m.role,
      content: typeof m.content === "string" ? m.content : JSON.stringify(m.content),
    })),
  ];

  for (;;) {
    const res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${API_KEY}` },
      body: JSON.stringify({
        model: MODEL,
        messages: apiMessages,
        tools: OPENAI_TOOLS,
        response_format: { type: "json_object" },
      }),
    });
    if (!res.ok) throw new Error(`OpenAI error ${res.status}: ${await res.text()}`);
    const data = await res.json();
    const choice = data.choices[0];

    if (choice.finish_reason !== "tool_calls") {
      return choice.message.content ?? "";
    }

    apiMessages.push(choice.message);
    for (const tc of choice.message.tool_calls ?? []) {
      const args = JSON.parse(tc.function.arguments || "{}");
      const result = await executeTool(tc.function.name, args);
      apiMessages.push({ role: "tool", tool_call_id: tc.id, content: result });
    }
  }
}

// ---------------------------------------------------------------------------
// Google Gemini agentic loop
// ---------------------------------------------------------------------------
async function runGoogle(messages, executeTool) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${API_KEY}`;
  const contents = messages.map((m) => ({
    role: m.role === "assistant" ? "model" : "user",
    parts: [{ text: typeof m.content === "string" ? m.content : JSON.stringify(m.content) }],
  }));

  for (;;) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        system_instruction: { parts: [{ text: SYSTEM_PROMPT }] },
        contents,
        tools: GOOGLE_TOOLS,
      }),
    });
    if (!res.ok) throw new Error(`Google error ${res.status}: ${await res.text()}`);
    const data = await res.json();
    const candidate = data.candidates[0];
    const parts = candidate.content.parts ?? [];

    const fnCalls = parts.filter((p) => p.functionCall);
    if (!fnCalls.length) {
      return parts.find((p) => p.text)?.text ?? "";
    }

    contents.push({ role: "model", parts });

    const responseParts = [];
    for (const p of fnCalls) {
      const result = await executeTool(p.functionCall.name, p.functionCall.args ?? {});
      responseParts.push({ functionResponse: { name: p.functionCall.name, response: { output: result } } });
    }
    contents.push({ role: "user", parts: responseParts });
  }
}

// ---------------------------------------------------------------------------
// Parse the LLM text response into { message, actions }
// ---------------------------------------------------------------------------
export function parseResponse(text) {
  try {
    const cleaned = text.replace(/^```(?:json)?\n?/i, "").replace(/\n?```$/i, "").trim();
    return JSON.parse(cleaned);
  } catch {
    return { message: text, actions: [] };
  }
}
