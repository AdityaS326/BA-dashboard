// backend/controllers/momController.js
import { callGroq } from "../utils/groq.js";

export async function generateMom(req, res) {
  const {
    title = "Meeting",
    date,
    attendees = "Not specified",
    facilitator = "Aditya S",
    objective = "",
    transcript = "",
  } = req.body;

  const meetingDate =
    date ||
    new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" });

  const prompt = `Generate a formal Minutes of Meeting (MOM) document.

Meeting Title : ${title}
Date          : ${meetingDate}
Attendees     : ${attendees}
Facilitator   : ${facilitator}
Objective     : ${objective || "General meeting"}
Transcript    : ${transcript || "(No transcript â€” generate realistic meeting content for this context)"}

Include:
1. Meeting header with all details
2. Agenda items discussed
3. Key Discussion Points (4â€“6 specific points)
4. Decisions Made (2â€“4 concrete decisions)
5. Action Items table (Action | Owner | Due Date | Priority)
6. Next Steps
7. Next Meeting suggestion
8. Sign-off section

Format as a clean, professional MOM. Use formal language.`;

  try {
    const text = await callGroq(
      prompt,
      "You are a professional business analyst who writes formal, well-structured Minutes of Meeting documents.",
      1200
    );
    res.json({ text });
  } catch (err) {
    console.error("[momController]", err.message);
    res.status(502).json({ error: err.message });
  }
}

