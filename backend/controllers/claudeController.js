// backend/controllers/claudeController.js
import { callGroq } from "../utils/groq.js";

export async function chat(req, res) {
  const { prompt, system } = req.body;
  if (!prompt) return res.status(400).json({ error: "prompt is required" });

  try {
    const text = await callGroq(
      prompt,
      system ||
        "You are a helpful, knowledgeable general assistant. Help with any professional task: writing, analysis, BRDs, user stories, emails, process flows, risk analysis, and more. Be concise and practical.",
      1500
    );
    res.json({ text });
  } catch (err) {
    console.error("[claudeController]", err.message);
    res.status(502).json({ error: err.message });
  }
}

