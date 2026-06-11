// backend/routes/report.js
import { Router } from "express";
import { requireGroqKey } from "../middleware/index.js";
import { generateReport } from "../controllers/reportController.js";

const router = Router();
router.post("/", requireGroqKey, generateReport);
export default router;

