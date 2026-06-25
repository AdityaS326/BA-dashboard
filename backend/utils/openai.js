// backend/utils/openai.js
// Thin wrapper around the OpenAI Chat Completions API.

import { config } from "../config/index.js";

/**
 * Call the OpenAI Chat Completions API.
 * @param {string} prompt      — User message
 * @param {string} [system]    — System prompt
 * @param {number} [maxTokens=1500]
 */
export async function callOpenAI(prompt, system, maxTokens = 1500) {
  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type":  "application/json",
      "Authorization": `Bearer ${config.openaiApiKey}`,
    },
    body: JSON.stringify({
      model:      config.openaiModel,
      max_tokens: maxTokens,
      messages: [
        { role: "system", content: system || "You are a helpful, professional assistant." },
        { role: "user",   content: prompt },
      ],
    }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`OpenAI API error ${response.status}: ${err}`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content || "";
}
