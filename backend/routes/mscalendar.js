// backend/routes/mscalendar.js
import { Router } from "express";
import { getEvents } from "../controllers/msCalendarController.js";

const router = Router();
router.get("/events", getEvents);
export default router;

