// backend/controllers/reportController.js
import { callAI } from "../utils/aiRouter.js";

export async function generateReport(req, res) {
  const {
    name      = "Team Member",
    dept      = "Technology",
    role      = "",
    project   = "",
    startTime,
    endTime,
    context   = "",
    standupHistory = [],
    provider  = "groq",
  } = req.body;

  const fmt = (dt) => dt
    ? new Date(dt).toLocaleString("en-IN", { day: "2-digit", month: "long", year: "numeric", hour: "2-digit", minute: "2-digit" })
    : null;

  const period = fmt(startTime) && fmt(endTime)
    ? `${fmt(startTime)} to ${fmt(endTime)}`
    : fmt(startTime) || fmt(endTime)
      || new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" });

  const roleLabel  = role?.trim()    || "Team Member";
  const projectStr = project?.trim() ? ` on the **${project.trim()}** project` : "";

  let standupBlock = "";
  if (standupHistory?.length) {
    const entries = standupHistory.slice(0, 10).map((s, i) => {
      const lines = [];
      if (s.date)     lines.push(`  Date: ${s.date}`);
      if (s.done)     lines.push(`  Completed: ${s.done}`);
      if (s.today)    lines.push(`  Planned next: ${s.today}`);
      if (s.blockers) lines.push(`  Blockers: ${s.blockers}`);
      return `Entry ${i + 1}:\n${lines.join("\n")}`;
    }).join("\n\n");
    standupBlock = `\nSTAND-UP HISTORY (actual work entries — use these as primary evidence):\n${entries}\n`;
  }

  const contextBlock = context?.trim()
    ? `\nADDITIONAL CONTEXT (provided by ${name}):\n${context.trim()}\n`
    : "";

  const hasRealData = standupBlock || contextBlock;

  const prompt =
    `Generate a professional weekly work report for ${name}, ${roleLabel}${projectStr} at ESDS Software Solution Pvt. Ltd.\n` +
    `Report period: ${period}\n` +
    `Department: ${dept}\n` +
    standupBlock +
    contextBlock +
    `\nINSTRUCTIONS:\n` +
    (hasRealData
      ? `- The stand-up history and context above are the PRIMARY source of truth. The report must reflect this SPECIFIC work, not generic descriptions.\n` +
        `- Reference actual tasks, features, tickets, or items mentioned in the stand-up entries.\n`
      : `- No stand-up history was provided. Generate a realistic report based on the role (${roleLabel}) and department.\n`) +
    `- Determine the appropriate report sections based on what was actually done. Do NOT use a rigid fixed template.\n` +
    `- Typical sections to consider (include what's relevant, skip what isn't): ` +
    `Executive Summary, Key Accomplishments, Work In Progress, Meetings & Collaboration, Challenges & How They Were Handled, ` +
    `Learning & Research, Metrics or Deliverables (if applicable), Plan for Next Week.\n` +
    `- Tailor the language and focus to the person's role: a ${roleLabel} uses different vocabulary than a Developer or PM.\n` +
    `- Write in formal third-person tone. Be specific — this should read like THIS person's week, not a generic template.\n` +
    `- If stand-up data mentions a specific feature, bug, requirement, or meeting — reference it by name.\n` +
    `- End with a clear "Next Week" plan derived from the "Planned next" entries in the stand-up history.`;

  const systemPrompt =
    `You are a professional report writer for enterprise software teams at ESDS Software Solution Pvt. Ltd. ` +
    `You write formal weekly work reports that are specific, insightful, and grounded in actual work evidence. ` +
    `You never produce generic boilerplate — every report should sound like it was written about this specific person's specific week. ` +
    `Determine the report structure dynamically based on the content provided. Use formal third-person tone throughout.`;

  try {
    const text = await callAI(prompt, systemPrompt, 5000, provider);
    res.json({ text, provider, meta: { name, dept, period, role: roleLabel } });
  } catch (err) {
    console.error("[reportController]", err.message);
    res.status(502).json({ error: err.message });
  }
}
