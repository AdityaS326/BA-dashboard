// backend/controllers/reportController.js
import { callAI } from "../utils/aiRouter.js";

export async function generateReport(req, res) {
  const {
    name = "Mahesh Beesu",
    dept = "Technology",
    startTime,
    endTime,
    context = "",
    provider = "groq",
  } = req.body;

  const fmt = (dt) => dt
    ? new Date(dt).toLocaleString("en-IN", { day: "2-digit", month: "long", year: "numeric", hour: "2-digit", minute: "2-digit" })
    : null;

  const period = fmt(startTime) && fmt(endTime)
    ? `${fmt(startTime)} to ${fmt(endTime)}`
    : fmt(startTime) || fmt(endTime)
      || new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" });

  const contextSection = context?.trim()
    ? `\nKey activities / context provided by the employee:\n${context.trim()}\n`
    : "";

  const prompt = `Generate a formal 9-section Weekly Work Report.

Employee: ${name}
Department: ${dept}
Report Period: ${period}
${contextSection}
Sections required:
SECTION 1 - EXECUTIVE SUMMARY (2-3 paragraphs)
SECTION 2 - KEY ACCOMPLISHMENTS (5-8 bullet points)
SECTION 3 - PROJECTS & TASKS IN PROGRESS
SECTION 4 - MEETINGS & COLLABORATIONS
SECTION 5 - RESEARCH & LEARNING
SECTION 6 - CLIENT & STAKEHOLDER INTERACTIONS
SECTION 7 - CHALLENGES & RESOLUTIONS
SECTION 8 - METRICS & DELIVERABLES
SECTION 9 - PLAN FOR NEXT WEEK (5-7 priorities)

Context: ${name} is a Business Analyst at ESDS Software Solution Pvt. Ltd.
Use formal third-person tone throughout. Be specific and realistic.`;

  try {
    const text = await callAI(
      prompt,
      "You are a professional report writer generating formal weekly work reports for enterprise software companies. Use formal third-person tone. Structure each section with clear headings, bullet points, and complete sentences.",
      4000,
      provider
    );
    res.json({ text, provider, meta: { name, dept, period } });
  } catch (err) {
    console.error("[reportController]", err.message);
    res.status(502).json({ error: err.message });
  }
}
