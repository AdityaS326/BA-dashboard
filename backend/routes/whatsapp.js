// backend/routes/whatsapp.js
import { Router } from 'express';
import {
  getStatus,
  sendMessage,
  getConversations,
  getMessages,
  webhookVerify,
  webhookReceive,
  getProfile,
} from '../controllers/whatsappController.js';

const router = Router();

router.get('/status',          getStatus);
router.get('/profile',         getProfile);
router.get('/conversations',   getConversations);
router.get('/messages/:phone', getMessages);
router.post('/send',           sendMessage);
router.get('/webhook',         webhookVerify);
router.post('/webhook',        webhookReceive);

export default router;
