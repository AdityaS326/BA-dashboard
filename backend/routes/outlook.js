// backend/routes/outlook.js
import { Router } from "express";
import { getEmails, getEmailBody, draftReply, sendEmail } from "../controllers/outlookController.js";

const router = Router();
router.get("/emails",          getEmails);
router.get("/emails/:id/body", getEmailBody);
router.post("/draft",          draftReply);
router.post("/send",           sendEmail);
export default router;

