# Office Agent — Script Lab AI Assistant

A local agentic server that connects Microsoft Word and PowerPoint (via [Script Lab](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/script-lab-overview)) to any major LLM provider. Chat with your documents, generate flowcharts, proofread with tracked changes, and run custom workflows — all from within Office.

---

## Architecture

```
┌─────────────────────────────┐         ┌───────────────────────────────┐
│  Script Lab (Word / PPT)    │  WebSocket  │  office-agent-server (Node)   │
│  snippets/word-agent.yaml   │◄──────────►│  server/index.js              │
│  snippets/powerpoint-agent  │            │  ├── llm.js  (Anthropic/OpenAI/│
│                             │            │  │           Google)            │
│  • Chat UI                  │            │  ├── db.js   (SQLite workflows) │
│  • Workflow picker          │            │  ├── flowchart.js               │
│  • Slash commands           │            │  └── proofread.js               │
└─────────────────────────────┘            └───────────────────────────────┘
                                                         │
                                                         ▼
                                                  LLM API (cloud)
```

The server handles all LLM calls and generates Office.js code server-side (guaranteed styling). The Script Lab snippets are thin WebSocket clients — they send context and receive executable code back.

---

## Quick Start

### 1. Install and configure the server

```bash
cd server
npm install
cp .env.example .env
# Edit .env with your API key and preferred model
```

**.env** options:

| Variable | Description | Example |
|---|---|---|
| `LLM_PROVIDER` | `anthropic`, `openai`, or `google` | `anthropic` |
| `LLM_API_KEY` | API key for the chosen provider | `sk-ant-...` |
| `LLM_MODEL` | Model name (provider-specific) | `claude-sonnet-4-6` |
| `PORT` | Server port (default `3579`) | `3579` |

**Model examples:**

| Provider | Models |
|---|---|
| Anthropic | `claude-sonnet-4-6`, `claude-opus-4-7`, `claude-haiku-4-5-20251001` |
| OpenAI | `gpt-4o`, `gpt-4o-mini`, `o3` |
| Google | `gemini-2.5-pro`, `gemini-2.5-flash` |

### 2. Start the server

```bash
npm start          # production
npm run dev        # auto-restart on file changes (Node 18+)
```

The server prints its URL and WebSocket path on startup:

```
Office Agent Server running on http://localhost:3579
  Provider : anthropic
  Model    : claude-sonnet-4-6
  WS path  : ws://localhost:3579/ws
```

### 3. Load the Script Lab snippets

1. Open Word or PowerPoint.
2. Open **Script Lab** (Insert → Script Lab → Code).
3. Click the hamburger menu → **Import** → paste the contents of the relevant YAML file:
   - `snippets/word-agent.yaml` — for Word
   - `snippets/powerpoint-agent.yaml` — for PowerPoint
4. Click **Run** in Script Lab.

The task pane will connect to the server automatically and show a chat interface.

---

## Features

### Chat

Type any message in the task pane. The agent reads your document context (selected text or body) and responds with both a chat reply and Office.js actions executed live in your document or slide.

### Workflows

Click a workflow button to activate a structured task. Workflows inject detailed instructions into the LLM prompt and constrain what it returns. The server then generates guaranteed, consistently-styled Office.js code from the structured data.

Built-in workflows:

| Workflow | What it does |
|---|---|
| **Summarize selection** | Condenses selected Word text into 3–5 bullet points and inserts them below the selection |
| **Rewrite (formal tone)** | Rewrites selected text in formal, professional language and replaces it in-document |
| **Extract action items** | Pulls every task or commitment from selected text into a numbered list |
| **Explain in plain English** | Explains selected text in simple language (chat only, no document change) |
| **Generate flowchart** | Parses steps from selected/body text and draws a styled flowchart on the active PowerPoint slide |
| **Generate key questions** | Produces 5–8 critical questions about the selected text |
| **Proofread document** | Checks the entire document for grammar, clarity, style, and consistency; applies changes as tracked edits with comments |

### Slash commands

Type a slash command in the chat input for quick workflow access:

| Command | Equivalent to |
|---|---|
| `/generate-flowchart` | Generate flowchart workflow |
| `/proofread` | Proofread document workflow |

### Custom workflows

Create your own workflows via the Manage Workflows panel (gear icon in the task pane):

1. Click **New workflow**.
2. Give it a title and write a prompt in Markdown.
3. The prompt is injected verbatim into the LLM message when the workflow is activated.

Custom workflows support the same structured-data return format as built-ins. To have the server generate flowchart code, instruct the LLM to return `flowchart_steps`. To trigger tracked proofreading, instruct it to return `proofreading_changes`.

---

## Flowchart Generation

When the **Generate flowchart** workflow runs, the LLM returns a `flowchart_steps` array (structured JSON, no code). The server's `flowchart.js` renders it with enforced styling:

- White fill, black 1.5pt border, black 12pt text
- Shape types: `FlowChartTerminator` (start/end, rendered as rectangles), `FlowChartProcess` (steps), `FlowChartDecision` (diamond, taller)
- Curved lead lines to the right of each shape
- Bold reference numerals (X02, X04, X06 … or user-specified) at the end of each lead line
- Straight arrow connectors between steps (requires PowerPointApi 1.6+; falls back gracefully)
- Text centering applied in a separate sync pass

The flowchart is drawn on the currently selected slide, falling back to slide 1.

---

## Proofreading

The **Proofread document** workflow uses the full document body as context. The LLM returns a `proofreading_changes` array (find / replace / reason). The server's `proofread.js`:

1. Enables `Word.ChangeTrackingMode.trackAll`
2. Searches for each `find` phrase
3. Inserts an inline comment with the reason
4. Replaces the text, so the change appears as a tracked deletion + insertion

Reviewers can then Accept or Reject each change through Word's Review pane.

---

## Server API

The server exposes both a WebSocket endpoint and an HTTP fallback.

### WebSocket (`ws://localhost:3579/ws`)

| Message type | Direction | Description |
|---|---|---|
| `register` | client→server | Start a session. Send `{ type, sessionId?, app, meta? }` |
| `registered` | server→client | Confirms session. Returns `{ type, sessionId }` |
| `message` | client→server | Chat turn. Send `{ type, id, content, app, context?, workflow? }` |
| `message` | server→client | LLM reply. Returns `{ type, id, content }` |
| `execute` | server→client | Office.js code to run. Returns `{ type, target, code }` |
| `result` | client→server | Execution result. Send `{ type, success, output? }` |
| `workflow.list` | client→server | List workflows |
| `workflow.create` | client→server | Create a custom workflow |
| `workflow.update` | client→server | Update a custom workflow |
| `workflow.delete` | client→server | Delete a custom workflow |
| `workflow.hide` | client→server | Hide a workflow from the picker |
| `workflow.unhide` | client→server | Unhide a workflow |

### HTTP

| Endpoint | Method | Description |
|---|---|---|
| `/message` | POST | Single-turn chat (no WebSocket required) |
| `/session/:id` | DELETE | Clear conversation history |
| `/health` | GET | Server status, provider, model, session count |

---

## Project Structure

```
scriptlab/
├── server/
│   ├── index.js          # Express + WebSocket server, action dispatcher
│   ├── llm.js            # Model-agnostic LLM client (Anthropic / OpenAI / Google)
│   ├── db.js             # SQLite workflow store (better-sqlite3)
│   ├── flowchart.js      # Server-side PowerPoint flowchart code generator
│   ├── proofread.js      # Server-side Word tracked-change code generator
│   ├── workflows/
│   │   └── builtins.js   # Built-in workflow definitions (upserted on startup)
│   ├── package.json
│   └── .env.example
└── snippets/
    ├── word-agent.yaml       # Script Lab snippet for Word
    └── powerpoint-agent.yaml # Script Lab snippet for PowerPoint
```

---

## Requirements

- Node.js 18+
- Microsoft 365 (Word / PowerPoint desktop or online) with Script Lab installed
- An API key for Anthropic, OpenAI, or Google

Script Lab is a free Microsoft add-in: [aka.ms/getscriptlab](https://aka.ms/getscriptlab)
