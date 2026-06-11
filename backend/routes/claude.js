// backend/routes/claude.js
import { Router } from "express";
import { requireGroqKey } from "../middleware/index.js";
import { chat } from "../controllers/claudeController.js";

const router = Router();
router.post("/", requireGroqKey, chat);
export default router;

