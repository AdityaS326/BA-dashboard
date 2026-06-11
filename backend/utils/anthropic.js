// backend/utils/anthropic.js
// Thin wrapper around the Anthropic Messages API.

import { config } from "../config/index.js";

/**
 * Call the Anthropic Claude API.
 * @param {string}   prompt   — User message
 * @param {string}   [system] — System prompt
 * @param {number}   [maxTokens=1500]
 * @param {string[]} [tools]  — Optional tool definitions (e.g. web_search)
 */
export async function callClaude(prompt, system, maxTokens = 1500, tools = []) {
  const body = {
    model:      config.anthropicModel,
    max_tokens: maxTokens,
    system:     system || "You are a helpful, professional assistant.",
    messages:   [{ role: "user", content: prompt }],
  };
  if (tools.length > 0) body.tools = tools;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type":    "application/json",
      "x-api-key":       config.anthropicApiKey,
      "anthropic-version": config.anthropicVersion,
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Anthropic API error ${response.status}: ${err}`);
  }

  const data = await response.json();
  return (data.content || [])
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("");
}
