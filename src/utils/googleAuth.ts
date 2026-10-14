import { initializeApp, getApps, getApp } from 'firebase/app';
import {
  getAuth,
  signInWithPopup,
  GoogleAuthProvider,
  onAuthStateChanged,
  signOut,
  type User,
} from 'firebase/auth';
import defaultFirebaseConfig from '../../firebase-applet-config.json';

const activeConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY || defaultFirebaseConfig.apiKey,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || defaultFirebaseConfig.authDomain,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || defaultFirebaseConfig.projectId,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || defaultFirebaseConfig.storageBucket,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || defaultFirebaseConfig.messagingSenderId,
  appId: import.meta.env.VITE_FIREBASE_APP_ID || defaultFirebaseConfig.appId,
};

const app = getApps().length === 0 ? initializeApp(activeConfig) : getApp();
export const auth = getAuth(app);
export const firebaseProjectId = activeConfig.projectId;

export const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

const STORAGE_TOKEN_KEY = 'gs_google_access_token';
const STORAGE_TIME_KEY = 'gs_google_token_time';

let cachedAccessToken: string | null =
  typeof window !== 'undefined' ? localStorage.getItem(STORAGE_TOKEN_KEY) : null;

// Creates a fresh GoogleAuthProvider with prompt=select_account
// This ensures that Google ALWAYS lets the user choose which account to use
// (e.g. pocop703@gmail.com, personal, work, etc.) and allows typing ANY email!
export const createGoogleAuthProvider = () => {
  const provider = new GoogleAuthProvider();
  SCOPES.forEach((scope) => provider.addScope(scope));
  provider.setCustomParameters({
    prompt: 'select_account',
  });
  return provider;
};

export const initAuth = (
  onAuthSuccess?: (user: User, token: string | null) => void,
  onAuthFailure?: () => void
) => {
  return onAuthStateChanged(auth, async (user: User | null) => {
    if (user) {
      if (onAuthSuccess) onAuthSuccess(user, cachedAccessToken);
    } else {
      cachedAccessToken = null;
      if (typeof window !== 'undefined') {
        localStorage.removeItem(STORAGE_TOKEN_KEY);
        localStorage.removeItem(STORAGE_TIME_KEY);
      }
      if (onAuthFailure) onAuthFailure();
    }
  });
};

export const googleSignIn = async (): Promise<{ user: User; accessToken: string } | null> => {
  try {
    const provider = createGoogleAuthProvider();
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Falha ao obter token de acesso do Google.');
    }

    cachedAccessToken = credential.accessToken;
    if (typeof window !== 'undefined') {
      localStorage.setItem(STORAGE_TOKEN_KEY, credential.accessToken);
      localStorage.setItem(STORAGE_TIME_KEY, Date.now().toString());
    }
    return { user: result.user, accessToken: cachedAccessToken };
  } catch (error: any) {
    console.error('Sign-in error:', error);
    if (error?.code === 'auth/unauthorized-domain' || error?.message?.includes('auth/unauthorized-domain')) {
      const currentHost = typeof window !== 'undefined' ? window.location.hostname : 'seu-dominio';
      const customErr: any = new Error(
        `O domínio '${currentHost}' precisa ser adicionado aos Domínios Autorizados no Firebase Console para permitir o login com o Google.`
      );
      customErr.code = 'auth/unauthorized-domain';
      customErr.domain = currentHost;
      customErr.projectId = activeConfig.projectId;
      throw customErr;
    }
    throw error;
  }
};

export const switchGoogleAccount = async (): Promise<{ user: User; accessToken: string } | null> => {
  try {
    await signOut(auth);
  } catch (e) {
    console.warn('Error signing out before switch account:', e);
  }
  cachedAccessToken = null;
  if (typeof window !== 'undefined') {
    localStorage.removeItem(STORAGE_TOKEN_KEY);
    localStorage.removeItem(STORAGE_TIME_KEY);
  }
  return googleSignIn();
};

export const getAccessToken = async (): Promise<string | null> => {
  if (cachedAccessToken) return cachedAccessToken;
  if (typeof window !== 'undefined') {
    return localStorage.getItem(STORAGE_TOKEN_KEY);
  }
  return null;
};

export const setAccessToken = (token: string | null) => {
  cachedAccessToken = token;
  if (typeof window !== 'undefined') {
    if (token) {
      localStorage.setItem(STORAGE_TOKEN_KEY, token);
      localStorage.setItem(STORAGE_TIME_KEY, Date.now().toString());
    } else {
      localStorage.removeItem(STORAGE_TOKEN_KEY);
      localStorage.removeItem(STORAGE_TIME_KEY);
    }
  }
};

export const logoutGoogle = async () => {
  await signOut(auth);
  cachedAccessToken = null;
  if (typeof window !== 'undefined') {
    localStorage.removeItem(STORAGE_TOKEN_KEY);
    localStorage.removeItem(STORAGE_TIME_KEY);
  }
};
