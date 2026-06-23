// backend/routes/mom.js
import { Router } from "express";
import { requireGroqKey } from "../middleware/index.js";
import { generateMom } from "../controllers/momController.js";

const router = Router();
router.post("/", requireGroqKey, generateMom);
export default router;

