// backend/routes/sharepoint.js
import { Router } from "express";
import { testConnection, exportReport, oauthCallback, refreshToken, listFiles, listSites, listSiteFiles } from "../controllers/sharepointController.js";

const router = Router();
router.post("/test",        testConnection);
router.post("/export",      exportReport);
router.get("/callback",     oauthCallback);
router.post("/refresh",     refreshToken);
router.get("/files",        listFiles);
router.get("/sites",        listSites);
router.get("/site-files",   listSiteFiles);
export default router;
