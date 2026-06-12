// backend/utils/aiRouter.js
// Unified AI call — routes to Groq, OpenAI, or Anthropic based on `provider`.
// Falls back to Groq if an API key is missing for the requested provider.

import { callGroq }   from "./groq.js";
import { callOpenAI } from "./openai.js";
import { callClaude } from "./anthropic.js";
import { config }     from "../config/index.js";

/**
 * @param {string} prompt
 * @param {string} system
 * @param {number} maxTokens
 * @param {"groq"|"openai"|"anthropic"} provider
 */
export async function callAI(prompt, system, maxTokens = 4000, provider = "groq") {
  switch (provider) {
    case "openai":
      if (!config.openaiApiKey) throw new Error("OPENAI_API_KEY is not set. Add it to your .env file.");
      return callOpenAI(prompt, system, maxTokens);

    case "anthropic":
      if (!config.anthropicApiKey) throw new Error("ANTHROPIC_API_KEY is not set. Add it to your .env file.");
      return callClaude(prompt, system, maxTokens);

    default: // "groq" or anything unknown
      if (!config.groqApiKey) throw new Error("GROQ_API_KEY is not set. Add it to your .env file.");
      return callGroq(prompt, system, maxTokens);
  }
}
