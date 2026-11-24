// Firebase: sign-in and the cloud database. Everything the rest of the app
// needs from Firebase goes through the functions exported here.
import { initializeApp } from 'https://www.gstatic.com/firebasejs/13.4.0/firebase-app.js';
import {
  getAuth, onAuthStateChanged, GoogleAuthProvider, GithubAuthProvider,
  signInWithPopup, signInWithRedirect, getRedirectResult, signOut, linkWithCredential, linkWithPopup,
} from 'https://www.gstatic.com/firebasejs/13.4.0/firebase-auth.js';
import {
  initializeFirestore, getFirestore, persistentLocalCache, persistentMultipleTabManager,
  doc, collection, onSnapshot, setDoc, deleteDoc, getDoc, getDocs, writeBatch,
} from 'https://www.gstatic.com/firebasejs/13.4.0/firebase-firestore.js';

const firebaseConfig = {
  apiKey: 'AIzaSyALO8g4Jb2nSHwOoOnG6ENIzHlghDF4r9k',
  authDomain: 'biledger.firebaseapp.com',
  projectId: 'biledger',
  storageBucket: 'biledger.firebasestorage.app',
  messagingSenderId: '4093066963',
  appId: '1:4093066963:web:d42ef239e6df261a069e26',
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
auth.languageCode = 'zh-TW';
let db;
try {
  db = initializeFirestore(app, {
    ignoreUndefinedProperties: true,
    localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
  });
} catch (e) {
  db = getFirestore(app);
}

const PROVIDERS = { google: () => new GoogleAuthProvider(), github: () => new GithubAuthProvider() };
const LABEL = { 'google.com': 'Google', 'github.com': 'GitHub' };
let pending = null; // credential waiting to be linked after signing in with the other provider

export const providerLabel = id => LABEL[id] || id;

export function onUser(cb) {
  getRedirectResult(auth).catch(e => cb(null, explain(e)));
  return onAuthStateChanged(auth, async user => {
    if (user && pending) {
      try { await linkWithCredential(user, pending.cred); } catch (e) { /* already linked */ }
      pending = null;
    }
    cb(user);
  });
}

export async function login(kind) {
  const provider = PROVIDERS[kind]();
  try {
    await signInWithPopup(auth, provider);
  } catch (e) {
    if (e.code === 'auth/popup-blocked' || e.code === 'auth/operation-not-supported-in-this-environment') {
      return signInWithRedirect(auth, provider);
    }
    if (e.code === 'auth/account-exists-with-different-credential') {
      const C = kind === 'github' ? GithubAuthProvider : GoogleAuthProvider;
      const cred = C.credentialFromError(e);
      if (cred) pending = { cred, kind };
    }
    throw explain(e);
  }
}

export async function linkProvider(kind) {
  try { await linkWithPopup(auth.currentUser, PROVIDERS[kind]()); }
  catch (e) { throw explain(e); }
}

export const logout = () => signOut(auth);

function explain(e) {
  const map = {
    'auth/account-exists-with-different-credential': '這個 email 之前用另一種方式登入過。請改按另一個登入按鈕，登入後會自動把兩種方式連在一起。',
    'auth/popup-closed-by-user': '登入視窗被關掉了，請再試一次。',
    'auth/cancelled-popup-request': '登入視窗被關掉了，請再試一次。',
    'auth/unauthorized-domain': '這個網址還沒加進 Firebase 的「授權網域」。',
    'auth/operation-not-allowed': '這種登入方式還沒在 Firebase 開啟。',
    'auth/credential-already-in-use': '這個帳號已經連結到另一個使用者了。',
    'auth/provider-already-linked': '這種登入方式已經連結過了。',
    'auth/network-request-failed': '連不上網路，請檢查網路後再試一次。',
  };
  const err = new Error(map[e.code] || `登入失敗（${e.code || e.message}）`);
  err.code = e.code;
  return err;
}

/* ---------- database: users/{uid}/meta/config, expenses/{id}, moves/{id}, receipts/{id} ---------- */
let uid = null;
const col = name => collection(db, 'users', uid, name);
const cfgRef = () => doc(db, 'users', uid, 'meta', 'config');

export function subscribe(userId, { onConfig, onExpenses, onMoves, onSync }) {
  uid = userId;
  const offs = [
    onSnapshot(cfgRef(), { includeMetadataChanges: true }, s => {
      onSync?.(!s.metadata.hasPendingWrites, s.metadata.fromCache);
      onConfig(s.exists() ? s.data() : null, s.metadata.fromCache);
    }, e => onConfig(null, false, e)),
    onSnapshot(col('expenses'), s => onExpenses(s.docs.map(d => d.data()))),
    onSnapshot(col('moves'), s => onMoves(s.docs.map(d => d.data()))),
  ];
  return () => offs.forEach(f => f());
}

export const saveConfig = cfg => setDoc(cfgRef(), cfg);
export const putExpense = e => setDoc(doc(col('expenses'), e.id), e);
export const delExpense = id => Promise.all([deleteDoc(doc(col('expenses'), id)), deleteDoc(doc(col('receipts'), id))]);
export const putMove = m => setDoc(doc(col('moves'), m.id), m);
export const delMove = id => deleteDoc(doc(col('moves'), id));
export const putReceipt = (id, r) => setDoc(doc(col('receipts'), id), r);
export const delReceipt = id => deleteDoc(doc(col('receipts'), id));
export async function getReceipt(id) { const s = await getDoc(doc(col('receipts'), id)); return s.exists() ? s.data() : null; }

// Replace everything (used by Excel import). Receipts are kept.
export async function replaceAll(cfg, expenses, moves) {
  const ops = [];
  for (const name of ['expenses', 'moves']) {
    const snap = await getDocs(col(name));
    snap.forEach(d => ops.push(b => b.delete(d.ref)));
  }
  expenses.forEach(e => ops.push(b => b.set(doc(col('expenses'), e.id), e)));
  moves.forEach(m => ops.push(b => b.set(doc(col('moves'), m.id), m)));
  ops.push(b => b.set(cfgRef(), cfg));
  for (let i = 0; i < ops.length; i += 400) {
    const b = writeBatch(db);
    ops.slice(i, i + 400).forEach(f => f(b));
    await b.commit();
  }
}
