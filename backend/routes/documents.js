import { Router } from "express";
import multer     from "multer";
import { uploadDocument } from "../controllers/documentsController.js";

const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 25 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ok = [
      "application/pdf",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/msword",
    ].includes(file.mimetype);
    ok ? cb(null, true) : cb(new Error("Only PDF and Word (.doc/.docx) files are supported."));
  },
});

const router = Router();
router.post("/upload", upload.single("file"), uploadDocument);
export default router;
