// backend/config/index.js
// Central config — reads from process.env (populated by dotenv)

export const config = {
  port:             parseInt(process.env.PORT || "3000"),
  nodeEnv:          process.env.NODE_ENV || "development",
  groqApiKey: process.env.GROQ_API_KEY || "",
  groqModel:  process.env.GROQ_MODEL  || "llama-3.3-70b-versatile",
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
