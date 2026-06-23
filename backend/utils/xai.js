// backend/utils/xai.js
// Wrapper around the xAI (Grok) Chat Completions API — OpenAI-compatible format.

import { config } from "../config/index.js";

export async function callXAI(prompt, system, maxTokens = 1500) {
  const response = await fetch("https://api.x.ai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type":  "application/json",
      "Authorization": `Bearer ${config.xaiApiKey}`,
    },
    body: JSON.stringify({
      model:      config.xaiModel,
      max_tokens: maxTokens,
      messages: [
        { role: "system", content: system || "You are a helpful, professional assistant." },
        { role: "user",   content: prompt },
      ],
    }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`xAI API error ${response.status}: ${err}`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content || "";
}
