import { Request, Response, NextFunction } from 'express';
import { adminAuth } from '../lib/firebase-admin.ts';
import { DecodedIdToken } from 'firebase-admin/auth';

export interface AuthRequest extends Request {
  user?: DecodedIdToken;
}

export const requireAuth = async (
  req: AuthRequest,
  res: Response,
  next: NextFunction
): Promise<any> => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'Unauthorized: Missing token' });
  }

  const token = authHeader.split('Bearer ')[1];
  try {
    const decodedToken = await adminAuth.verifyIdToken(token);
    
    // Check Email Allowlist
    const allowedEmails = process.env.VITE_ALLOWED_EMAILS || process.env.ALLOWED_EMAILS || process.env.ADMIN_EMAILS;
    if (allowedEmails && allowedEmails.trim() !== '') {
      const emailList = allowedEmails.split(',').map(e => e.trim().toLowerCase());
      if (decodedToken.email && !emailList.includes(decodedToken.email.toLowerCase())) {
        return res.status(403).json({ error: 'Forbidden: Your email is not authorized' });
      }
    }

    req.user = decodedToken;
    next();
  } catch (error) {
    console.error('Error verifying Firebase ID token:', error);
    return res.status(401).json({ error: 'Unauthorized: Invalid token' });
  }
};
