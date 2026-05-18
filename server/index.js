import "dotenv/config";
import express from "express";
import { createServer } from "http";
import { WebSocketServer, WebSocket } from "ws";
import { v4 as uuidv4 } from "uuid";
import { chat, parseResponse } from "./llm.js";
import { db, buildWorkflowStore } from "./db.js";
import { buildFlowchartCode } from "./flowchart.js";
import { buildProofreadCode } from "./proofread.js";

const PORT = process.env.PORT || 3579;

// ---------------------------------------------------------------------------
// Express
// ---------------------------------------------------------------------------
const app = express();
app.use(express.json());
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

// ---------------------------------------------------------------------------
// In-memory state
// ---------------------------------------------------------------------------
const conversations = new Map(); // sessionId → [{role, content}]
const clients = new Map();       // sessionId → { ws, app, meta }

// ---------------------------------------------------------------------------
// WebSocket server
// ---------------------------------------------------------------------------
const httpServer = createServer(app);
const wss = new WebSocketServer({ server: httpServer, path: "/ws" });

wss.on("connection", (ws) => {
  let sessionId = null;

  ws.on("message", async (raw) => {
    let msg;
    try { msg = JSON.parse(raw.toString()); }
    catch { return ws.send(JSON.stringify({ type: "error", content: "Invalid JSON" })); }

    // ── register ─────────────────────────────────────────────────────────────
    if (msg.type === "register") {
      sessionId = msg.sessionId || uuidv4();
      clients.set(sessionId, { ws, app: msg.app, meta: msg.meta || {} });
      if (!conversations.has(sessionId)) conversations.set(sessionId, []);
      ws.send(JSON.stringify({ type: "registered", sessionId }));
      console.log(`[WS] registered  session=${sessionId}  app=${msg.app}`);
      return;
    }

    if (!sessionId) {
      return ws.send(JSON.stringify({ type: "error", content: "Send a register message first" }));
    }

    // ── workflow.list ─────────────────────────────────────────────────────────
    if (msg.type === "workflow.list") {
      const includeHidden = msg.includeHidden === true;
      const hidden = new Set(
        db.prepare("SELECT workflow_id FROM hidden_workflows").all().map((r) => r.workflow_id)
      );
      const rows = db
        .prepare("SELECT * FROM workflows ORDER BY is_system ASC, title ASC")
        .all();

      const workflows = rows
        .filter((w) => includeHidden || !hidden.has(w.id))
        .map((w) => ({
          id: w.id,
          title: w.title,
          type: w.type,
          practice: w.practice,
          is_system: !!w.is_system,
          hidden: hidden.has(w.id),
          prompt_md: includeHidden ? w.prompt_md : undefined,
          prompt_md_preview: w.prompt_md ? w.prompt_md.slice(0, 200) : null,
        }));

      ws.send(JSON.stringify({ type: "workflow.list.ok", id: msg.id, workflows }));
      return;
    }

    // ── workflow.create ───────────────────────────────────────────────────────
    if (msg.type === "workflow.create") {
      const id = uuidv4();
      db.prepare(
        `INSERT INTO workflows (id, title, type, prompt_md, practice, is_system)
         VALUES (?, ?, ?, ?, ?, 0)`
      ).run(id, msg.title, msg.wfType || "assistant", msg.prompt_md || null, msg.practice || null);
      ws.send(JSON.stringify({ type: "workflow.create.ok", id: msg.id, workflow_id: id }));
      return;
    }

    // ── workflow.update ───────────────────────────────────────────────────────
    if (msg.type === "workflow.update") {
      const wf = db.prepare("SELECT * FROM workflows WHERE id = ?").get(msg.workflow_id);
      if (!wf || wf.is_system) {
        return ws.send(JSON.stringify({ type: "error", id: msg.id, content: "Cannot update a built-in workflow" }));
      }
      const p = msg.patch || {};
      db.prepare(
        `UPDATE workflows SET title=?, prompt_md=?, practice=? WHERE id=?`
      ).run(
        p.title ?? wf.title,
        p.prompt_md ?? wf.prompt_md,
        "practice" in p ? p.practice : wf.practice,
        msg.workflow_id
      );
      ws.send(JSON.stringify({ type: "workflow.update.ok", id: msg.id, workflow_id: msg.workflow_id }));
      return;
    }

    // ── workflow.delete ───────────────────────────────────────────────────────
    if (msg.type === "workflow.delete") {
      const wf = db.prepare("SELECT * FROM workflows WHERE id = ?").get(msg.workflow_id);
      if (!wf || wf.is_system) {
        return ws.send(JSON.stringify({ type: "error", id: msg.id, content: "Cannot delete a built-in workflow" }));
      }
      db.prepare("DELETE FROM workflows WHERE id = ?").run(msg.workflow_id);
      ws.send(JSON.stringify({ type: "workflow.delete.ok", id: msg.id, workflow_id: msg.workflow_id }));
      return;
    }

    // ── workflow.hide / unhide ────────────────────────────────────────────────
    if (msg.type === "workflow.hide") {
      db.prepare("INSERT OR IGNORE INTO hidden_workflows (workflow_id) VALUES (?)").run(msg.workflow_id);
      ws.send(JSON.stringify({ type: "workflow.hide.ok", id: msg.id, workflow_id: msg.workflow_id }));
      return;
    }

    if (msg.type === "workflow.unhide") {
      db.prepare("DELETE FROM hidden_workflows WHERE workflow_id = ?").run(msg.workflow_id);
      ws.send(JSON.stringify({ type: "workflow.unhide.ok", id: msg.id, workflow_id: msg.workflow_id }));
      return;
    }

    // ── message ───────────────────────────────────────────────────────────────
    if (msg.type === "message") {
      const history = conversations.get(sessionId);

      const workflowStore = buildWorkflowStore();

      let userText = msg.content;
      if (msg.workflow?.id) {
        const wf = workflowStore.get(msg.workflow.id);
        if (wf) {
          // Inject the full workflow instructions directly — more reliable than requiring a read_workflow tool call
          userText = `## Active Workflow: ${wf.title}\n\nFollow these instructions exactly for this turn:\n\n${wf.prompt_md}\n\n## User Message\n\n${userText}`;
        }
      }

      const userContent = buildUserContent({ ...msg, content: userText });
      history.push({ role: "user", content: userContent });
      let agentReply;
      try {
        const rawText = await chat(history, { workflowStore });
        agentReply = parseResponse(rawText);
      } catch (err) {
        console.error("[LLM error]", err.message);
        return ws.send(JSON.stringify({ type: "error", content: err.message }));
      }

      // Flowchart: LLM returns structured data; server generates guaranteed code
      if (agentReply.flowchart_steps?.length) {
        agentReply.actions = [{ target: "powerpoint", code: buildFlowchartCode(agentReply.flowchart_steps) }];
        console.log(`[flowchart] ${agentReply.flowchart_steps.length} steps → code generated server-side`);
      } else if (msg.workflow?.id === "builtin-flowchart-from-steps" && agentReply.actions?.length) {
        // LLM ignored the workflow format and generated code directly — wrap it with style enforcement
        console.warn("[flowchart] LLM returned actions instead of flowchart_steps — applying style wrapper");
        agentReply.actions = agentReply.actions.map(a =>
          a.target === "powerpoint" && a.code ? { ...a, code: wrapWithFlowchartStyle(a.code) } : a
        );
      }

      // Proofreading: LLM returns structured changes; server generates tracked-edit code
      if (agentReply.proofreading_changes?.length) {
        agentReply.actions = [{ target: "word", code: buildProofreadCode(agentReply.proofreading_changes) }];
        console.log(`[proofread] ${agentReply.proofreading_changes.length} change(s) → tracked-edit code generated server-side`);
      }

      history.push({ role: "assistant", content: JSON.stringify(agentReply) });
      ws.send(JSON.stringify({ type: "message", id: msg.id, content: agentReply.message }));

      if (agentReply.actions?.length) dispatchActions(agentReply.actions, sessionId);
      return;
    }

    // ── result ────────────────────────────────────────────────────────────────
    if (msg.type === "result") {
      console.log(`[result] session=${sessionId}  success=${msg.success}  output=${msg.output}`);
      return;
    }
  });

  ws.on("close", () => {
    if (sessionId) {
      clients.delete(sessionId);
      console.log(`[WS] disconnected  session=${sessionId}`);
    }
  });

  ws.on("error", (err) => console.error("[WS error]", err.message));
});

// ---------------------------------------------------------------------------
// HTTP fallback
// ---------------------------------------------------------------------------
app.post("/message", async (req, res) => {
  const { sessionId: sid, message, app: appName, context, workflow } = req.body;
  const sessionId = sid || uuidv4();
  if (!conversations.has(sessionId)) conversations.set(sessionId, []);
  const history = conversations.get(sessionId);

  const workflowStore = buildWorkflowStore();

  let userText = message;
  if (workflow?.id) {
    const wf = workflowStore.get(workflow.id);
    if (wf) {
      userText = `## Active Workflow: ${wf.title}\n\nFollow these instructions exactly for this turn:\n\n${wf.prompt_md}\n\n## User Message\n\n${userText}`;
    }
  }

  const userContent = buildUserContent({ app: appName, content: userText, context });
  history.push({ role: "user", content: userContent });
  let agentReply;
  try {
    const rawText = await chat(history, { workflowStore });
    agentReply = parseResponse(rawText);
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }

  if (agentReply.flowchart_steps?.length) {
    agentReply.actions = [{ target: "powerpoint", code: buildFlowchartCode(agentReply.flowchart_steps) }];
  }

  if (agentReply.proofreading_changes?.length) {
    agentReply.actions = [{ target: "word", code: buildProofreadCode(agentReply.proofreading_changes) }];
  }

  history.push({ role: "assistant", content: JSON.stringify(agentReply) });
  dispatchActions(agentReply.actions || [], sessionId);
  res.json({ sessionId, message: agentReply.message, actions: agentReply.actions || [] });
});

app.delete("/session/:id", (req, res) => {
  conversations.delete(req.params.id);
  res.json({ ok: true });
});

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    provider: process.env.LLM_PROVIDER,
    model: process.env.LLM_MODEL,
    sessions: conversations.size,
    connected: clients.size,
  });
});

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function buildUserContent(msg) {
  const ctx = msg.context ? `\n\nContext: ${JSON.stringify(msg.context)}` : "";
  return `[App: ${msg.app || "unknown"}] ${msg.content}${ctx}`;
}

function dispatchActions(actions, senderSessionId) {
  for (const action of actions) {
    if (!action.code) continue;
    const targetApp = action.target;
    for (const [sid, client] of clients.entries()) {
      const appMatches = targetApp === "all" || client.app?.toLowerCase() === targetApp?.toLowerCase();
      if (!appMatches || client.ws.readyState !== WebSocket.OPEN) continue;
      client.ws.send(
        JSON.stringify({ type: "execute", target: targetApp, code: action.code, originSession: senderSessionId })
      );
      console.log(`[dispatch] → session=${sid}  app=${client.app}  target=${targetApp}`);
    }
  }
}

// ---------------------------------------------------------------------------
// Flowchart style enforcement wrapper
// Used as a fallback when the LLM generates raw PowerPoint code instead of
// returning flowchart_steps. Appends a second PowerPoint.run that applies
// white fill and black text to every shape on the active slide.
// ---------------------------------------------------------------------------
function wrapWithFlowchartStyle(code) {
  return `(async () => {
  ${code}

  await PowerPoint.run(async (ctx) => {
    let slide;
    try {
      const sel = ctx.presentation.getSelectedSlides?.();
      if (sel) {
        sel.load("items");
        await ctx.sync();
        if (sel.items?.length > 0) slide = sel.items[0];
      }
    } catch (_) {}
    if (!slide) {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      slide = slides.items[0];
    }
    slide.shapes.load("items/type");
    await ctx.sync();
    for (const shape of slide.shapes.items) {
      try { shape.fill.setSolidColor("FFFFFF"); } catch (_) {}
      try { shape.lineFormat.color = "000000"; shape.lineFormat.weight = 1.5; } catch (_) {}
      try { shape.textFrame.textRange.font.color = "000000"; } catch (_) {}
    }
    await ctx.sync();
  });
})()`;
}

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------
httpServer.listen(PORT, () => {
  console.log(`Office Agent Server running on http://localhost:${PORT}`);
  console.log(`  Provider : ${process.env.LLM_PROVIDER}`);
  console.log(`  Model    : ${process.env.LLM_MODEL}`);
  console.log(`  WS path  : ws://localhost:${PORT}/ws`);
});
