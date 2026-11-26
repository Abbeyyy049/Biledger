// Sends the daily reminder notifications. Run every hour by .github/workflows/notify.yml.
// For each person who turned notifications on, at the hour they chose (in their time zone),
// it checks every ledger they belong to and sends one notification per ledger that has something due.
import { remindersFor, messageFor, localNow } from './reminders.mjs';

const ROLES = ['owner', 'editor', 'commenter', 'viewer'];
const APP_URL = 'https://abbeyyy049.github.io/Biledger/';
const GONE = ['messaging/registration-token-not-registered', 'messaging/invalid-registration-token', 'messaging/invalid-argument'];

export async function run({ db, send, now = new Date(), FieldPath, FieldValue, log = console.log }) {
  const users = await db.collection('users').listDocuments();
  let sent = 0;
  for (const u of users) {
    const ref = u.collection('meta').doc('push'), snap = await ref.get();
    const p = snap.exists ? snap.data() : null;
    const tokens = Object.keys(p?.tokens || {});
    if (!p?.on || !tokens.length) continue;
    const { date, hour } = localNow(now, p.tz);
    if (hour !== (p.hour ?? 9) || p.lastSent === date) continue;
    const ledgers = await db.collection('ledgers').where(new FieldPath('members', u.id), 'in', ROLES).get();
    const messages = [];
    for (const l of ledgers.docs) {
      if (p.ledgers?.[l.id] === false) continue;
      const [cfg, ex, mv] = await Promise.all([l.ref.collection('meta').doc('config').get(), l.ref.collection('expenses').get(), l.ref.collection('moves').get()]);
      if (!cfg.exists) continue;
      const r = remindersFor({ cfg: cfg.data(), expenses: ex.docs.map(d => d.data()), moves: mv.docs.map(d => d.data()) }, date, p.types || {});
      const m = messageFor(l.data().name, r, cfg.data().currency);
      if (m) messages.push({ ...m, lid: l.id });
    }
    const dead = new Set();
    for (const m of messages) {
      const res = await send(tokens, { title: m.title, body: m.body, link: `${APP_URL}?ledger=${m.lid}#home`, lid: m.lid });
      res.forEach((x, i) => { if (!x.ok && GONE.includes(x.code)) dead.add(tokens[i]); });
      sent += res.filter(x => x.ok).length;
    }
    await ref.update({ lastSent: date, lastCount: messages.length });
    // phones that uninstalled the app or turned notifications off
    for (const t of dead) await ref.update(new FieldPath('tokens', t), FieldValue.delete());
    log(`user ${u.id.slice(0, 6)}…: ${messages.length} notification(s) for ${date} ${hour}:00 (${p.tz}); removed ${dead.size} expired device(s)`);
  }
  return sent;
}

// Run from the command line / GitHub Actions
if (import.meta.url === `file://${process.argv[1]}`) {
  const { initializeApp, cert } = await import('firebase-admin/app');
  const { getFirestore, FieldPath, FieldValue } = await import('firebase-admin/firestore');
  const { getMessaging } = await import('firebase-admin/messaging');
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT;
  if (!raw && !process.env.FIRESTORE_EMULATOR_HOST) { console.error('FIREBASE_SERVICE_ACCOUNT is not set (add it under Settings → Secrets and variables → Actions).'); process.exit(1); }
  initializeApp(raw ? { credential: cert(JSON.parse(raw)) } : { projectId: 'biledger' });
  const messaging = getMessaging();
  const send = async (tokens, m) => {
    const r = await messaging.sendEachForMulticast({
      tokens,
      data: { lid: m.lid, link: m.link },
      webpush: { notification: { title: m.title, body: m.body, icon: `${APP_URL}icons/icon-192.png` }, fcmOptions: { link: m.link } },
    });
    return r.responses.map(x => ({ ok: x.success, code: x.error?.code }));
  };
  const n = await run({ db: getFirestore(), send, FieldPath, FieldValue });
  console.log(`done: ${n} notification(s) delivered`);
}
