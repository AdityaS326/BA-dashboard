// backend/routes/whatsapp.js
import { Router } from 'express';
import {
  initWhatsApp, getStatus, getChats,
  refreshChats, logoutWhatsApp,
} from '../controllers/whatsappController.js';

const router = Router();

router.post('/init',    initWhatsApp);
router.get('/status',   getStatus);
router.get('/chats',    getChats);
router.post('/refresh', refreshChats);
router.post('/logout',  logoutWhatsApp);

export default router;
