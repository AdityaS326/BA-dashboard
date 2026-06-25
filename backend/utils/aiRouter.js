// backend/utils/aiRouter.js
// Unified AI call — routes to Groq, OpenAI, or Anthropic based on `provider`.
// Falls back to Groq if an API key is missing for the requested provider.

import { callGroq }   from "./groq.js";
import { callOllama } from "./ollama.js";
import { callClaude } from "./anthropic.js";
import { config }     from "../config/index.js";

/**
 * @param {string} prompt
 * @param {string} system
 * @param {number} maxTokens
 * @param {"groq"|"ollama"|"anthropic"} provider
 */
export async function callAI(prompt, system, maxTokens = 4000, provider = "groq") {
  switch (provider) {
    case "ollama":
      return callOllama(prompt, system, maxTokens);

    case "anthropic":
      if (config.anthropicApiKey) return callClaude(prompt, system, maxTokens);
      console.warn("[aiRouter] ANTHROPIC_API_KEY not set, falling back to Groq");
      // fall through to groq

    default: // "groq" or fallback
      if (!config.groqApiKey) throw new Error("GROQ_API_KEY is not set. Add it to your .env file.");
      return callGroq(prompt, system, maxTokens);
  }
}
