// backend/routes/teams.js
import { Router } from "express";
import { getMeetings, generateTeamsMOM } from "../controllers/teamsController.js";

const router = Router();
router.get("/meetings",  getMeetings);
router.post("/mom",      generateTeamsMOM);
export default router;

