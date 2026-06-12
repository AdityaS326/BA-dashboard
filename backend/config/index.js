// backend/config/index.js
// Central config — reads from process.env (populated by dotenv)

export const config = {
  port:    parseInt(process.env.PORT || "3000"),
  nodeEnv: process.env.NODE_ENV || "development",

  // Groq (LLaMA 3.3 70B — default)
  groqApiKey: process.env.GROQ_API_KEY || "",
  groqModel:  process.env.GROQ_MODEL   || "llama-3.3-70b-versatile",

  // OpenAI (ChatGPT)
  openaiApiKey: process.env.OPENAI_API_KEY || "",
  openaiModel:  process.env.OPENAI_MODEL   || "gpt-4o",

  // Anthropic (Claude)
  anthropicApiKey:  process.env.ANTHROPIC_API_KEY     || "",
  anthropicModel:   process.env.ANTHROPIC_MODEL        || "claude-sonnet-4-6",
  anthropicVersion: process.env.ANTHROPIC_API_VERSION  || "2023-06-01",

  sharepoint: {
    tenantId:     process.env.SP_TENANT_ID     || "",
    clientId:     process.env.SP_CLIENT_ID     || "",
    clientSecret: process.env.SP_CLIENT_SECRET || "",
  },
};

export function requireEnv(key) {
  if (!process.env[key]) {
    console.warn(`⚠  Missing environment variable: ${key}`);
  }
}
