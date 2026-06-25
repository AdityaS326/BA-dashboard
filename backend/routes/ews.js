import { Router } from "express";
import { getMeetings, getEmails, getEmailBody, sendEmail, discoverEWS, createMeeting } from "../controllers/ewsController.js";

const router = Router();

router.post("/meetings",       getMeetings);
router.post("/emails",         getEmails);
router.post("/email-body",     getEmailBody);
router.post("/send-email",     sendEmail);
router.post("/create-meeting", createMeeting);
router.get("/discover",        discoverEWS);

export default router;
