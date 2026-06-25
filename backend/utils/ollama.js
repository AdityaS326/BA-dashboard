// backend/utils/ollama.js
// Wrapper around the local Ollama API (OpenAI-compatible endpoint).

import { config } from "../config/index.js";

export async function callOllama(prompt, system, maxTokens = 4000) {
  const baseUrl = config.ollamaBaseUrl || "http://localhost:11434";
  const model   = config.ollamaModel   || "llama3.2";

  let response;
  try {
    response = await fetch(`${baseUrl}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model,
        stream: false,
        options: { num_predict: maxTokens },
        messages: [
          { role: "system", content: system || "You are a helpful, professional assistant." },
          { role: "user",   content: prompt },
        ],
      }),
    });
  } catch {
    throw new Error(
      `Ollama is not running at ${baseUrl}. ` +
      `Start it with: ollama serve — then try again, or switch to Groq.`
    );
  }

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    if (response.status === 404) {
      throw new Error(
        `Ollama model "${model}" not found. Pull it first: ollama pull ${model}`
      );
    }
    throw new Error(`Ollama error ${response.status}: ${body}`);
  }

  const data = await response.json();
  return data.message?.content || "";
}
