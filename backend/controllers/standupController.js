// backend/controllers/standupController.js
import { callAI } from "../utils/aiRouter.js";

const SYSTEM_STANDUP =
  "You are a stand-up update generator who deeply understands software team workflows. " +
  "Before writing, reason through: " +
  "(1) What did the person actually accomplish yesterday — completed, in-review, handed off, or ongoing? " +
  "(2) Does today's plan naturally follow from yesterday — for example, a developer who pushed a fix yesterday might be writing tests or moving to the next ticket today; " +
  "a BA who finished requirements yesterday might be getting stakeholder sign-off today. " +
  "(3) What does this work mean at this stage of the project — is it a milestone, a dependency, a risk item? " +
  "Then write ONE paragraph that shows you understood this work progression, not just their notes. " +
  "Reveal insight: if today logically follows yesterday, show that cause-and-effect naturally. " +
  "Never use bullet points, headers, or numbered lists. Never sound robotic or corporate. " +
  "Use contractions. Show the story of the work.";

const SYSTEM_QA =
  "You help team members answer manager questions during stand-up meetings. " +
  "Before answering, understand the full work context — what stage things are at, what decisions were made, what the risks are. " +
  "Give a confident, specific answer that shows the person understands their own work. " +
  "Sound like a real person speaking, not a report. Never start with the word 'I'.";

function roleProfile(role) {
  if (!role) return null;
  const r = role.toLowerCase();

  if (/\bdev(eloper|elop)?\b|engineer|programmer|frontend|backend|fullstack/.test(r))
    return {
      label: "Developer",
      vocab: "code, commits, PRs, bug fixes, branches, reviews, deployments, unit tests, APIs, modules, refactoring",
      example: "Yesterday I pushed the fix for the auth token issue and got it reviewed. Today I'm working on the dashboard API endpoint — should have it ready for QA by end of day.",
    };

  if (/\bpm\b|project.?manager|program.?manager/.test(r))
    return {
      label: "Project Manager",
      vocab: "sprint planning, timelines, milestones, stakeholder updates, risk tracking, resource allocation, blockers, delivery dates, team coordination",
      example: "Yesterday I ran the sprint review and updated the delivery timeline with the client. Today I'm following up on the two open risks and syncing with the dev lead on the Q3 scope.",
    };

  if (/\bba\b|business.?analyst|analyst|solution.?architect/.test(r))
    return {
      label: "Business Analyst",
      vocab: "requirements, user stories, acceptance criteria, BRD, FRD, stakeholder mapping, gap analysis, process flows, sign-offs",
      example: "Yesterday I finished the user story mapping for the checkout flow and got stakeholder sign-off. Today I'm refining the acceptance criteria with the product owner and starting the gap analysis.",
    };

  if (/\bqa\b|test|quality/.test(r))
    return {
      label: "QA Engineer",
      vocab: "test cases, bug reports, regression tests, test coverage, automation scripts, defect triage, release validation",
      example: "Yesterday I completed regression testing for the payment module and logged three bugs. Today I'm working on automating the login flow tests and will be verifying the fixes from yesterday's build.",
    };

  if (/design|ux|ui/.test(r))
    return {
      label: "Designer",
      vocab: "wireframes, mockups, user flows, prototypes, design reviews, Figma, accessibility, component library",
      example: "Yesterday I finalised the onboarding wireframes and shared them for feedback. Today I'm iterating on the dashboard layout based on the comments and working on the mobile breakpoints.",
    };

  if (/manager|lead|head|director/.test(r))
    return {
      label: "Team Lead / Manager",
      vocab: "team check-ins, code reviews, planning, priorities, blockers, mentoring, delivery tracking",
      example: "Yesterday I reviewed two PRs and had a planning session with the team. Today I'm unblocking the integration issue and syncing with the product manager on next sprint priorities.",
    };

  return { label: role, vocab: `${role}-specific tasks and terminology`, example: null };
}

export async function generateStandup(req, res) {
  const {
    done     = "",
    today    = "",
    blockers = "",
    tone     = "casual",
    name     = "Team Member",
    role     = "",
    project  = "",
    provider = "groq",
  } = req.body;

  if (!done.trim() && !today.trim())
    return res.status(400).json({ error: "Provide at least yesterday or today input." });

  const date = new Date().toLocaleDateString("en-IN", {
    weekday: "long", day: "2-digit", month: "long", year: "numeric",
  });

  const rp          = roleProfile(role);
  const projectLine = project ? ` on the ${project} project` : "";
  const toneGuide   = tone === "professional"
    ? "Polished and professional, but still warm and human."
    : "Casual and friendly — like talking to teammates.";

  const roleSection = rp
    ? `\n${name} is a ${rp.label}${projectLine}.\n` +
      `Role vocabulary to use naturally: ${rp.vocab}.\n` +
      (rp.example ? `How a ${rp.label} sounds: "${rp.example}"\n` : "")
    : "";

  const hasYesterday = done.trim().length > 0;
  const hasToday     = today.trim().length > 0;
  const workFlow = hasYesterday && hasToday
    ? `Understand the work stage: given that ${name} did "${done}" yesterday and plans "${today}" today, ` +
      `determine whether today is a continuation, the next phase, or a new track — then express that relationship naturally.\n`
    : "";

  const prompt =
    `Date: ${date}.\n` +
    roleSection +
    (hasYesterday ? `Yesterday's work: ${done}\n` : "") +
    (hasToday     ? `Today's plan: ${today}\n`     : "") +
    (blockers?.trim() ? `Blockers: ${blockers}\n`  : "") +
    workFlow +
    `\nTone: ${toneGuide}\n\n` +
    `Output rules:\n` +
    `- ONE single natural paragraph — no lists, no headers, no bullet points.\n` +
    `- Sound like ${name} speaking out loud to their team, not writing a report.\n` +
    `- Show understanding of the work: if today follows logically from yesterday, make that clear naturally.\n` +
    `- Use ${rp ? rp.label + "-specific" : "role-appropriate"} language woven into real sentences.\n` +
    `- ONLY mention blockers if they were explicitly provided above — do NOT invent or guess any.\n` +
    `- Use contractions: I'm, I've, didn't, we're, it's.\n` +
    `- 4–6 sentences. Do NOT start with the word "I".\n` +
    `- Output ONLY the paragraph. No preamble, no label, no sign-off.`;

  try {
    const text = await callAI(prompt, SYSTEM_STANDUP, 520, provider);
    res.json({ text });
  } catch (err) {
    console.error("[standupController]", err.message);
    res.status(502).json({ error: err.message });
  }
}

export async function standupQA(req, res) {
  const {
    question    = "",
    done        = "",
    today       = "",
    blockers    = "",
    standupText = "",
    name        = "Team Member",
    role        = "",
    provider    = "groq",
  } = req.body;

  if (!question.trim()) return res.status(400).json({ error: "Question is required." });

  const rp = roleProfile(role);
  const roleHint = rp ? ` (${rp.label})` : "";

  const prompt =
    `${name}${roleHint} stand-up context:\n` +
    `- Yesterday: ${done || "Not specified"}\n` +
    `- Today: ${today || "Not specified"}\n` +
    `- Blockers: ${blockers || "None"}\n` +
    `- Full stand-up: ${standupText || "(not generated)"}\n\n` +
    `Manager's question: "${question}"\n\n` +
    `Write a confident, natural response ${name} would say out loud. ` +
    (rp ? `Use ${rp.label}-appropriate language. ` : "") +
    `2–4 sentences. Be specific. Sound like a real person, not a report.`;

  try {
    const answer = await callAI(prompt, SYSTEM_QA, 250, provider);
    res.json({ answer });
  } catch (err) {
    console.error("[standupQA]", err.message);
    res.status(502).json({ error: err.message });
  }
}
