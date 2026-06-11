// backend/controllers/standupController.js
import { callGroq } from "../utils/groq.js";

export async function standupQA(req, res) {
  const {
    question    = "",
    done        = "",
    today       = "",
    blockers    = "None",
    standupText = "",
    name        = "Aditya S",
  } = req.body;

  if (!question.trim()) return res.status(400).json({ error: "Question is required." });

  const prompt = `You are helping ${name} (a Business Analyst) respond to their manager's question during a standup.

Standup context:
- Yesterday: ${done || "Not specified"}
- Today: ${today || "Not specified"}
- Blockers: ${blockers || "None"}
- Full standup update: ${standupText || "Not generated yet"}

Manager's question: "${question}"

Write a confident, clear, first-person response that ${name} should give. Keep it to 2–4 sentences. Be direct and professional. Address the question specifically using the standup context.`;

  try {
    const answer = await callGroq(
      prompt,
      `You help Business Analysts respond confidently to manager questions in standup meetings. Sound like a real person — natural, first-person, concise. Never start with "I" as the literal first word.`,
      300
    );
    res.json({ answer });
  } catch (err) {
    console.error("[standupQA]", err.message);
    res.status(502).json({ error: err.message });
  }
}

export async function generateStandup(req, res) {
  const {
    done     = "",
    today    = "",
    blockers = "None",
    format   = "Bullet points (Teams/Slack)",
    name     = "Aditya S",
  } = req.body;

  const date = new Date().toLocaleDateString("en-IN", {
    weekday: "long", day: "2-digit", month: "long", year: "numeric",
  });

  const isEmail  = format.toLowerCase().includes("email");
  const isJira   = format.toLowerCase().includes("jira");
  const isFormal = format.toLowerCase().includes("formal") || isEmail;

  const prompt = `Write a ${isFormal ? "formal" : "casual, friendly"} daily stand-up for ${name}.

Date     : ${date}
Yesterday: ${done || "Not specified"}
Today    : ${today || "Not specified"}
Blockers : ${blockers || "None"}
Format   : ${format}

${isEmail
  ? "Write as a short email. Subject line first, then 3-4 sentences in first person."
  : isJira
    ? "Write as a JIRA comment with h3. headers. Keep it factual."
    : `Write in natural, human-sounding first-person language — like you're telling a teammate what you did.
3 short sections: ✅ Yesterday | 🔨 Today | 🚧 Blockers (omit blockers section if none).
Keep each section to 1-2 lines. Sound like a real person, not a report.`}`;

  try {
    const text = await callGroq(
      prompt,
      "You write concise, natural-sounding daily stand-up updates for a Business Analyst at a software company. Sound human, not corporate.",
      400
    );
    res.json({ text });
  } catch (err) {
    console.error("[standupController]", err.message);
    res.status(502).json({ error: err.message });
  }
}
