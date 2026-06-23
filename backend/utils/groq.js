// backend/utils/groq.js
// Wrapper around the Groq Chat Completions API (OpenAI-compatible).

import { config } from "../config/index.js";

export async function callGroq(prompt, system, maxTokens = 1500) {
  const response = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type":  "application/json",
      "Authorization": `Bearer ${config.groqApiKey}`,
    },
    body: JSON.stringify({
      model:      config.groqModel,
      max_tokens: maxTokens,
      messages: [
        { role: "system", content: system || "You are a helpful, professional assistant." },
        { role: "user",   content: prompt },
      ],
    }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Groq API error ${response.status}: ${err}`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content || "";
}
