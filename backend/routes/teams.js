// backend/routes/teams.js
import { Router } from "express";
import { getMeetings, generateTeamsMOM, getChats, getChatMessages } from "../controllers/teamsController.js";

const router = Router();
router.get("/meetings",                    getMeetings);
router.post("/mom",                        generateTeamsMOM);
router.get("/chats",                       getChats);
router.get("/chats/:chatId/messages",      getChatMessages);
export default router;

