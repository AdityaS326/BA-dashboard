// backend/routes/standup.js
import { Router } from "express";
import { requireGroqKey } from "../middleware/index.js";
import { generateStandup, standupQA } from "../controllers/standupController.js";

const router = Router();
router.post("/",   requireGroqKey, generateStandup);
router.post("/qa", requireGroqKey, standupQA);
export default router;

