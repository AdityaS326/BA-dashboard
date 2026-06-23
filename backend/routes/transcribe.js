// backend/routes/transcribe.js
import { Router }         from "express";
import multer             from "multer";
import { transcribeAudio } from "../controllers/transcribeController.js";

// 500 MB — large files are converted + chunked in the controller
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 500 * 1024 * 1024 } });

const router = Router();
router.post("/", upload.single("file"), transcribeAudio);
export default router;
