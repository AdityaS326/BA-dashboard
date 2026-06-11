// backend/routes/outlook.js
import { Router } from "express";
import { getEmails, draftReply, sendEmail } from "../controllers/outlookController.js";

const router = Router();
router.get("/emails",  getEmails);
router.post("/draft",  draftReply);
router.post("/send",   sendEmail);
export default router;

