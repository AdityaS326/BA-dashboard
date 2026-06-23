import { Router } from "express";
import { getMembers, saveMembers } from "../controllers/teamMembersController.js";

const router = Router();
router.get("/",  getMembers);
router.post("/", saveMembers);
export default router;
