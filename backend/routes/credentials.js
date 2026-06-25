// backend/routes/credentials.js
import { Router } from 'express';
import { getCredentials, saveCredentials } from '../controllers/credentialsController.js';

const router = Router();
router.get('/',  getCredentials);
router.post('/', saveCredentials);
export default router;
