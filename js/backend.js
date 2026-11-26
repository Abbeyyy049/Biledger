// Firebase: sign-in and the cloud database. Everything the rest of the app
// needs from Firebase goes through the functions exported here.
import { initializeApp } from 'https://www.gstatic.com/firebasejs/13.4.0/firebase-app.js';
import {
  getAuth, onAuthStateChanged, GoogleAuthProvider, GithubAuthProvider,
  signInWithPopup, signInWithRedirect, getRedirectResult, signOut, linkWithCredential, linkWithPopup,
  createUserWithEmailAndPassword, signInWithEmailAndPassword, sendPasswordResetEmail, EmailAuthProvider,
  sendEmailVerification, connectAuthEmulator,
} from 'https://www.gstatic.com/firebasejs/13.4.0/firebase-auth.js';
import {
  initializeFirestore, getFirestore, persistentLocalCache, persistentMultipleTabManager,
  doc, collection, onSnapshot, setDoc, deleteDoc, getDoc, getDocs, writeBatch, updateDoc, addDoc,
  query, where, FieldPath, deleteField, connectFirestoreEmulator,
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
// automated tests point the app at local Firebase emulators
if (window.__BILEDGER_EMULATOR__) {
  connectAuthEmulator(auth, 'http://127.0.0.1:9099', { disableWarnings: true });
  connectFirestoreEmulator(db, '127.0.0.1', 8080);
}

const PROVIDERS = { google: () => new GoogleAuthProvider(), github: () => new GithubAuthProvider() };
const LABEL = { 'google.com': 'Google', 'github.com': 'GitHub', password: 'Email' };
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

export async function emailLogin(email, pw, create) {
  try {
    const cred = await (create ? createUserWithEmailAndPassword : signInWithEmailAndPassword)(auth, email.trim(), pw);
    // new accounts get a confirmation email right away: shared ledgers only open for confirmed addresses
    if (create) await sendEmailVerification(cred.user).catch(() => {});
  } catch (e) { throw explain(e); }
}
export async function resetPassword(email) {
  try { await sendPasswordResetEmail(auth, email.trim()); } catch (e) { throw explain(e); }
}
// lets someone who signed in with Google/GitHub also sign in with email + password
export async function addPassword(email, pw) {
  try { await linkWithCredential(auth.currentUser, EmailAuthProvider.credential(email.trim(), pw)); }
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
    'auth/invalid-credential': 'Email 或密碼不對。',
    'auth/wrong-password': 'Email 或密碼不對。',
    'auth/user-not-found': '找不到這個 Email 的帳號，可以按「建立新帳號」。',
    'auth/invalid-email': 'Email 格式不對。',
    'auth/missing-password': '請輸入密碼。',
    'auth/weak-password': '密碼至少要 6 個字。',
    'auth/email-already-in-use': '這個 Email 已經有帳號了。如果之前是用 Google 或 GitHub 登入，請用那個方式登入，再到「設定 → 帳號與資料」加上密碼。',
    'auth/too-many-requests': '嘗試太多次了，請稍等一下再試。',
    'auth/requires-recent-login': '為了安全，請先登出再登入一次，然後再試。',
  };
  const err = new Error(map[e.code] || `沒有成功（${e.code || e.message}）`);
  err.code = e.code;
  return err;
}

export async function sendVerify() { try { await sendEmailVerification(auth.currentUser); } catch (e) { throw explain(e); } }
// after the person clicks the link in the email: refresh so the new status reaches the database rules
export async function refreshUser() { await auth.currentUser.reload(); await auth.currentUser.getIdToken(true); return auth.currentUser; }

/* ---------- ledgers ----------
   ledgers/{lid}                      name, owner, members {uid: role}, invites {emailKey: role}, names {uid: {name,email}}
   ledgers/{lid}/meta/config          groups, assets, categories, methods, jars, bills, currency, fiscal year
   ledgers/{lid}/expenses|moves|receipts|icons|comments/{id}
   users/{uid}/meta/prefs             this person's font and last opened ledger
   users/{uid}/...                    data from before ledgers existed (moved into a ledger once) */
export const ROLES = ['owner', 'editor', 'commenter', 'viewer'];
export const emailKey = e => String(e || '').trim().toLowerCase().replace(/\./g, ',');
export const keyEmail = k => String(k).replace(/,/g, '.');
let lid = null;
const col = name => collection(db, 'ledgers', lid, name);
const cfgRef = () => doc(db, 'ledgers', lid, 'meta', 'config');
const ledgerRef = id => doc(db, 'ledgers', id || lid);
const me = () => auth.currentUser;
const myName = () => ({ name: me().displayName || me().email || '', email: me().email || '' });

// ledgers I belong to, and invitations waiting for my (verified) email
export function watchLedgers(user, cb) {
  let mine = null, invites = [];
  const send = () => { if (mine) cb({ mine, invites }); };
  const offs = [onSnapshot(query(collection(db, 'ledgers'), where(new FieldPath('members', user.uid), 'in', ROLES)),
    s => { mine = s.docs.map(d => ({ id: d.id, ...d.data() })); send(); }, e => cb(null, e))];
  if (user.email && user.emailVerified) offs.push(onSnapshot(query(collection(db, 'ledgers'), where(new FieldPath('invites', emailKey(user.email)), 'in', ROLES.slice(1))),
    s => { invites = s.docs.map(d => ({ id: d.id, ...d.data() })); send(); }, () => {}));
  return () => offs.forEach(f => f());
}
export async function createLedger(name, cfg) {
  const ref = doc(collection(db, 'ledgers'));
  await setDoc(ref, { name, owner: me().uid, members: { [me().uid]: 'owner' }, invites: {}, names: { [me().uid]: myName() }, created: Date.now() });
  await setDoc(doc(db, 'ledgers', ref.id, 'meta', 'config'), cfg);
  return ref.id;
}
export const renameLedger = name => updateDoc(ledgerRef(), { name });
export const invite = (email, role) => updateDoc(ledgerRef(), new FieldPath('invites', emailKey(email)), role);
export const cancelInvite = key => updateDoc(ledgerRef(), new FieldPath('invites', key), deleteField());
export const setRole = (uid, role) => updateDoc(ledgerRef(), new FieldPath('members', uid), role);
export const removeMember = uid => updateDoc(ledgerRef(), new FieldPath('members', uid), deleteField(), new FieldPath('names', uid), deleteField());
export async function acceptInvite(id, role) {
  await updateDoc(ledgerRef(id), new FieldPath('members', me().uid), role, new FieldPath('invites', emailKey(me().email)), deleteField(), new FieldPath('names', me().uid), myName());
}
export const leaveLedger = id => updateDoc(ledgerRef(id), new FieldPath('members', me().uid), deleteField(), new FieldPath('names', me().uid), deleteField());
export async function deleteLedger() { await clearAll(); await deleteDoc(ledgerRef()); }

export async function getPrefs(uid) { const s = await getDoc(doc(db, 'users', uid, 'meta', 'prefs')).catch(() => null); return s?.exists() ? s.data() : {}; }
export const setPrefs = (uid, p) => setDoc(doc(db, 'users', uid, 'meta', 'prefs'), p, { merge: true }).catch(() => {});

// Data saved before ledgers existed lives under users/{uid}; copy it into a new ledger once.
export async function legacyData(uid) {
  const moved = await getDoc(doc(db, 'users', uid, 'meta', 'moved'));
  if (moved.exists()) return null;
  const cfg = await getDoc(doc(db, 'users', uid, 'meta', 'config'));
  return cfg.exists() ? cfg.data() : null;
}
export async function moveLegacy(uid, name, cfg) {
  const id = await createLedger(name, cfg);
  const prev = lid; lid = id;
  try {
    for (const name of ['expenses', 'moves', 'receipts', 'icons']) {
      const snap = await getDocs(collection(db, 'users', uid, name));
      for (let i = 0; i < snap.docs.length; i += 400) {
        const b = writeBatch(db);
        snap.docs.slice(i, i + 400).forEach(d => b.set(doc(col(name), d.id), d.data()));
        await b.commit();
      }
    }
    await setDoc(doc(db, 'users', uid, 'meta', 'moved'), { to: id, at: Date.now() });
  } finally { lid = prev; }
  return id;
}

/* ---------- records of the open ledger ---------- */

export function subscribe(ledgerId, { onLedger, onConfig, onExpenses, onMoves, onIcons, onComments, onSync }) {
  lid = ledgerId;
  const offs = [
    onSnapshot(cfgRef(), { includeMetadataChanges: true }, s => {
      onSync?.(!s.metadata.hasPendingWrites, s.metadata.fromCache);
      onConfig(s.exists() ? s.data() : null, s.metadata.fromCache);
    }, e => onConfig(null, false, e)),
    onSnapshot(col('expenses'), s => onExpenses(s.docs.map(d => d.data()))),
    onSnapshot(col('moves'), s => onMoves(s.docs.map(d => d.data()))),
    onSnapshot(col('icons'), s => onIcons?.(s.docs.map(d => ({ id: d.id, ...d.data() })))),
    onSnapshot(col('comments'), s => onComments?.(s.docs.map(d => ({ id: d.id, ...d.data() }))), () => {}),
    onSnapshot(ledgerRef(), s => onLedger?.(s.exists() ? { id: s.id, ...s.data() } : null), () => onLedger?.(null)),
  ];
  return () => offs.forEach(f => f());
}

export const saveConfig = cfg => setDoc(cfgRef(), cfg);
export const putExpense = e => setDoc(doc(col('expenses'), e.id), e);
export const delExpense = id => Promise.all([deleteDoc(doc(col('expenses'), id)), deleteDoc(doc(col('receipts'), id))]);
export const putMove = m => setDoc(doc(col('moves'), m.id), m);
export const delMove = id => deleteDoc(doc(col('moves'), id));
export const putIcon = (id, ic) => setDoc(doc(col('icons'), id), ic);
export const delIcon = id => deleteDoc(doc(col('icons'), id));
export const addComment = (exp, text) => addDoc(col('comments'), { exp, text, uid: me().uid, ...myName(), at: Date.now() });
export const delComment = id => deleteDoc(doc(col('comments'), id));
export const putReceipt = (id, r) => setDoc(doc(col('receipts'), id), r);
export const delReceipt = id => deleteDoc(doc(col('receipts'), id));
export async function getReceipt(id) { const s = await getDoc(doc(col('receipts'), id)); return s.exists() ? s.data() : null; }

// Delete every record and the settings of the signed-in person.
export async function clearAll() {
  const refs = [];
  for (const name of ['expenses', 'moves', 'receipts', 'icons', 'comments']) (await getDocs(col(name))).forEach(d => refs.push(d.ref));
  refs.push(cfgRef());
  for (let i = 0; i < refs.length; i += 400) { const b = writeBatch(db); refs.slice(i, i + 400).forEach(r => b.delete(r)); await b.commit(); }
}

// Replace everything (used by Excel import). Receipts are kept.
export async function replaceAll(cfg, expenses, moves) {
  const ops = [];
  for (const name of ['expenses', 'moves', 'comments']) {
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
