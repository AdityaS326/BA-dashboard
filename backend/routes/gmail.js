// backend/routes/gmail.js
import { Router } from 'express';
import {
  getAuthUrl, handleCallback, getStatus,
  getEmails, getEmailById, sendEmail, gmailLogout,
} from '../controllers/gmailController.js';

const router = Router();

router.get('/auth-url',  getAuthUrl);
router.get('/callback',  handleCallback);
router.get('/status',    getStatus);
router.get('/emails',    getEmails);
router.get('/email/:id', getEmailById);
router.post('/send',     sendEmail);
router.post('/logout',   gmailLogout);

export default router;
