// backend/controllers/reportController.js
import { callGroq } from "../utils/groq.js";

export async function generateReport(req, res) {
  const {
    name = "Aditya S",
    dept = "Technology",
    week,
    manager = "Igor (Product Owner)",
    source = "claude",
  } = req.body;

  const today =
    week ||
    new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" });

  const prompt = `Generate a formal 9-section Weekly Work Report.

Employee: ${name}
Department: ${dept}
Week Ending: ${today}
Manager: ${manager}
Source: ${source} AI conversation history

Sections required:
SECTION 1 â€” EXECUTIVE SUMMARY (2â€“3 paragraphs)
SECTION 2 â€” KEY ACCOMPLISHMENTS (5â€“8 bullet points)
SECTION 3 â€” PROJECTS & TASKS IN PROGRESS
SECTION 4 â€” MEETINGS & COLLABORATIONS
SECTION 5 â€” RESEARCH & LEARNING
SECTION 6 â€” CLIENT & STAKEHOLDER INTERACTIONS
SECTION 7 â€” CHALLENGES & RESOLUTIONS
SECTION 8 â€” METRICS & DELIVERABLES
SECTION 9 â€” PLAN FOR NEXT WEEK (5â€“7 priorities)

Context: ${name} is a System Analyst / Solution Architect at ESDS Software Solution Pvt. Ltd.
working on EEL (Enterprise Linux), GPOS Subscription Manager, BRD/FRD documentation, and solution architecture.
Use formal third-person tone throughout. Be specific and realistic.`;

  try {
    const text = await callGroq(
      prompt,
      "You are a professional report writer generating formal weekly work reports for enterprise software companies. Use formal third-person tone.",
      3000
    );
    res.json({ text, meta: { name, dept, week: today, manager } });
  } catch (err) {
    console.error("[reportController]", err.message);
    res.status(502).json({ error: err.message });
  }
}

