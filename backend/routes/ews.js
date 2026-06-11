import { Router } from "express";
import { getMeetings, getEmails, getEmailBody, sendEmail, discoverEWS } from "../controllers/ewsController.js";

const router = Router();

router.post("/meetings",    getMeetings);
router.post("/emails",      getEmails);
router.post("/email-body",  getEmailBody);
router.post("/send-email",  sendEmail);
router.get("/discover",     discoverEWS);

export default router;
