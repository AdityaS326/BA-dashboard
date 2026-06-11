// backend/routes/sharepoint.js
import { Router } from "express";
import { testConnection, exportReport, oauthCallback, refreshToken } from "../controllers/sharepointController.js";

const router = Router();
router.post("/test",     testConnection);
router.post("/export",   exportReport);
router.get("/callback",  oauthCallback);
router.post("/refresh",  refreshToken);
export default router;

