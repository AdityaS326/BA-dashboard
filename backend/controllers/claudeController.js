// backend/controllers/claudeController.js
import { callAI } from "../utils/aiRouter.js";

function buildSystemPrompt(projectName, userCtx = {}) {
  const { userName, userRole, recentWork, currentPlan } = userCtx;

  const projectLine  = projectName?.trim()  ? `\n- Current project: **${projectName.trim()}**`                        : "";
  const nameLine     = userName?.trim()     ? `\n- User: **${userName.trim()}**`                                       : "";
  const roleLine     = userRole?.trim()     ? `\n- Role: **${userRole.trim()}**`                                       : "";
  const workLine     = recentWork?.trim()   ? `\n- Recently working on: ${recentWork.trim().slice(0, 250)}`            : "";
  const planLine     = currentPlan?.trim()  ? `\n- Today's plan: ${currentPlan.trim().slice(0, 250)}`                  : "";

  const contextBlock = (projectLine || nameLine || roleLine || workLine || planLine)
    ? `\n\nUSER CONTEXT (use this to tailor every response):${projectLine}${nameLine}${roleLine}${workLine}${planLine}` +
      `\n\nWhen the user asks something, interpret it through their role and current work. ` +
      `If their question is vague or short, infer what they likely need based on their role and what they've been doing. ` +
      `Proactively connect your answer to their actual work where it adds value.`
    : "";

  return `You are an expert project assistant at ESDS Software Solution Pvt. Ltd., supporting the entire project team — developers, BAs, PMs, QAs, designers, and leads.${contextBlock}

RESPONSE STYLE:
- Always respond in well-structured, detailed format using sections, bullet points, and numbered lists where appropriate.
- Use clear headings (##) to separate sections.
- For documents (BRDs, user stories, RTMs, emails, plans): produce complete, professional, ready-to-use output — not summaries.
- For analysis tasks: lead with a summary, then provide detailed breakdown.
- For emails: use proper letter format with greeting, body paragraphs, and closing.
- Bold **key terms** for scannability.

DOMAIN EXPERTISE:
- Business Analysis: BRD, FRD, SRS, RTM, use cases, user stories, acceptance criteria, process flows, gap analysis, RACI, stakeholder mapping.
- Development: architecture decisions, technical documentation, API design, sprint planning, deployment checklists, code review guidance.
- Project Management: timelines, risk management, resource planning, status reporting, milestones.
- QA: test plans, test cases, bug triage, regression strategies, release validation.
- Tools: Jira, Confluence, SharePoint, Microsoft Exchange/Teams.`;
}

export async function chat(req, res) {
  const {
    prompt, system, provider = "groq", projectName = "",
    userName = "", userRole = "", recentWork = "", currentPlan = "",
  } = req.body;
  if (!prompt) return res.status(400).json({ error: "prompt is required" });

  try {
    const text = await callAI(
      prompt,
      system || buildSystemPrompt(projectName, { userName, userRole, recentWork, currentPlan }),
      4000,
      provider
    );
    res.json({ text, provider });
  } catch (err) {
    console.error("[claudeController]", err.message);
    res.status(502).json({ error: err.message });
  }
}
