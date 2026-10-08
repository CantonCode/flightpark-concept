/**
 * Flight Park memberships: proof of concept.
 *
 * The Google Sheet is the database, this script is the API, and Stripe takes payments.
 * No webhooks: the script creates each Stripe Checkout itself, confirms payment by asking
 * Stripe when the pilot comes back to the site, and a job every SYNC_MINUTES catches anyone
 * who closed the tab after paying.
 *
 * Script properties (Project Settings → Script properties):
 *   STRIPE_SECRET_KEY  your sk_test_… key while testing
 *   SITE_URL           where Stripe may send pilots back to; comma-separate several, e.g.
 *                      http://localhost:8766/, https://cantoncode.github.io/flightpark-concept/
 *   SEND_EMAILS        "yes" to email pilots a confirmation (optional)
 */

const VERSION = 'poc-9';
const TZ = 'Pacific/Auckland';
const TAB = { members: 'Members', companies: 'Companies', passes: 'Passes' };
const MEMBER_COLS = ['Member #', 'Added', 'First name', 'Last name', 'Email', 'Phone', 'NZHGPA PIN', 'Type', 'Pass',
  'Company', 'Starts', 'Expires', 'Show on list', 'Source', 'Stripe session', 'Amount paid', 'Notes'];
const FIRST_MEMBER_NO = 3000;
const SYNC_MINUTES = 15;  // backup check for pilots who pay and close the tab; Apps Script allows 1, 5, 10, 15 or 30
const SHOW_EXPIRED_DAYS = 30;  // expired pilots stay on the public list this long, marked expired
const MAX_FAILS = 5;
const LOCK_SECONDS = 15 * 60;
const PASSWORD_ERROR = 'Company password incorrect, try again or contact your manager for the correct password.';

/* ---------------------------------------------------------------- setup and sheet menu */

/** Run once from the editor. Creates the tabs with headers and sample settings. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  ensureTab_(ss, TAB.members, MEMBER_COLS, []);
  formatMembers_(membersTable_());
  ensureTab_(ss, TAB.companies, ['Company', 'Password', 'Active', 'Membership days'], [
    ['Coronet Peak Tandems', 'cpt-test', 'Yes', 365],
    ['Extreme Air', 'xair-test', 'Yes', 365],
    ['G-Force', 'gforce-test', 'Yes', 365],
    ['Infinity', 'infinity-test', 'Yes', 365],
    ['Paraventures', 'paraventures-test', 'Yes', 365],
    ['Skytrek', 'skytrek-test', 'Yes', 365],
  ]);
  ensureTab_(ss, TAB.passes, ['Pass', 'Price (NZD)', 'Label'], [
    ['day', 10, 'Day pass'], ['week', 20, 'Week pass'], ['month', 40, 'Month pass'], ['year', 80, 'Year pass'],
  ]);
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  SpreadsheetApp.getUi().alert('Setup done. Next: add your script properties, then deploy as a web app.');
}

/** Run once after adding STRIPE_SECRET_KEY, and again after changing SYNC_MINUTES. Checks Stripe for paid checkouts. */
function installSyncTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncPayments')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncPayments').timeBased().everyMinutes(SYNC_MINUTES).create();
}

/** Adds a "Flight Park" menu to the sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Flight Park')
    .addItem('Fill in missing member numbers', 'fillMemberNumbers')
    .addToUi();
}

/** Menu action: numbers any rows typed in by hand. Returning pilots keep their number. */
function fillMemberNumbers() {
  const n = withLock_(() => assignMissingNumbers_(membersTable_()));
  clearMembersCache_();
  SpreadsheetApp.getActive().toast(n ? n + ' row(s) numbered.' : 'Every row already has a member number.');
}

function ensureTab_(ss, name, headers, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sh.autoResizeColumns(1, headers.length);
  }
  return sh;
}

function formatMembers_(t) {
  const col = name => t.sh.getRange(1, t.map[name] + 1, t.sh.getMaxRows(), 1);
  col('Member #').setNumberFormat('0');
  col('Added').setNumberFormat('d mmm yyyy h:mm');
  col('Phone').setNumberFormat('@');
  col('NZHGPA PIN').setNumberFormat('@');
  col('Starts').setNumberFormat('d mmm yyyy');
  col('Expires').setNumberFormat('d mmm yyyy');
  col('Amount paid').setNumberFormat('$0.00');
  t.sh.getRange(2, t.map['Show on list'] + 1, t.sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
  t.sh.getRange(1, 1, 1, t.width).setFontWeight('bold');
}

/* ---------------------------------------------------------------- web app */

function doGet(e) {
  const p = (e && e.parameter) || {};
  return handle_(() => {
    if (p.action === 'ping') return { ok: true, version: VERSION };
    if (p.action === 'confirm') return confirmPayment_(p.session_id);
    return { ok: true, members: currentMembers_() };
  });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (_) {
    return json_({ ok: false, error: 'That request could not be read.' });
  }
  return handle_(() => {
    if (body.action === 'checkout') return createCheckout_(body);
    if (body.action === 'commercial') return commercialSignup_(body);
    if (body.action === 'confirm') return confirmPayment_(body.sessionId);
    return { ok: false, error: 'Unknown action.' };
  });
}

function handle_(fn) {
  try {
    return json_(fn());
  } catch (err) {
    console.error(err && err.stack || err);
    return json_({ ok: false, error: err.userMessage || 'Something went wrong on our side. Please try again.' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------------------------------------------------------- recreational: Stripe checkout */

function createCheckout_(b) {
  const d = validatePilot_(b, { pinRequired: false });
  if (d.error) return d;
  const pass = getPasses_()[String(b.pass || '').toLowerCase()];
  if (!pass) return fail_('Choose a pass length.', 'pass');
  const start = clean_(b.start);
  if (start) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start)) return fail_('Use a valid start date.', 'start');
    if (start < todayYmd_()) return fail_("The start date can't be in the past.", 'start');
  }
  const siteUrl = pickSiteUrl_(b.returnUrl);
  prop_('STRIPE_SECRET_KEY');

  // Returning pilots keep their number; new pilots get the next one, shown on the Stripe page.
  const memberNo = withLock_(() => {
    const t = membersTable_();
    return existingMemberNo_(t, d.email, d.pin, d.last) || nextMemberNo_(t);
  });

  const session = stripe_('post', '/v1/checkout/sessions', {
    'mode': 'payment',
    'customer_email': d.email,
    'line_items[0][quantity]': 1,
    'line_items[0][price_data][currency]': 'nzd',
    'line_items[0][price_data][unit_amount]': Math.round(pass.price * 100),
    'line_items[0][price_data][product_data][name]': 'Flight Park ' + pass.label,
    'line_items[0][price_data][product_data][description]': 'Member #' + memberNo + ' · ' + d.first + ' ' + d.last +
      (start && start > todayYmd_() ? ' · Starts ' + Utilities.formatDate(toDate_(start), TZ, 'd MMM yyyy') : ''),
    'payment_intent_data[description]': 'Flight Park ' + pass.label + ' · Member #' + memberNo,
    'payment_intent_data[metadata][member_no]': memberNo,
    'metadata[member_no]': memberNo,
    'metadata[first]': d.first,
    'metadata[last]': d.last,
    'metadata[phone]': d.phone,
    'metadata[pin]': d.pin,
    'metadata[pass]': pass.key,
    'metadata[start]': start,
    'success_url': siteUrl + '?paid={CHECKOUT_SESSION_ID}',
    'cancel_url': siteUrl + '#membership',
  });
  return { ok: true, url: session.url, memberNo };
}

/** The page the pilot came from, if it's one of the SITE_URL addresses; otherwise the first one. Never an address outside the list. */
function pickSiteUrl_(requested) {
  const norm = u => String(u || '').trim().replace(/[?#].*$/, '').replace(/index\.html$/, '');
  const allowed = prop_('SITE_URL').split(',').map(norm).filter(Boolean);
  const want = norm(requested);
  return allowed.find(u => u === want) || allowed[0];
}

/** Called when Stripe sends the pilot back with ?paid=cs_… */
function confirmPayment_(sessionId) {
  const id = clean_(sessionId);
  if (!/^cs_(test|live)_[A-Za-z0-9]+$/.test(id)) return fail_('That payment reference is not valid.');
  let session;
  try {
    session = stripe_('get', '/v1/checkout/sessions/' + encodeURIComponent(id));
  } catch (err) {
    if (err.status === 404) return fail_("We couldn't find that payment.");
    throw err;
  }
  if (session.payment_status !== 'paid') return { ok: false, status: session.payment_status, error: 'That payment has not gone through yet.' };
  return recordPaidSession_(session);
}

/** Time-driven fallback: adds any paid checkout from the last 2 days that is not in the sheet yet. */
function syncPayments() {
  const since = Math.floor(Date.now() / 1000) - 2 * 24 * 3600;
  const list = stripe_('get', '/v1/checkout/sessions?status=complete&limit=100&created%5Bgte%5D=' + since);
  const t = membersTable_();
  const known = new Set(t.rows.map(r => String(get_(t, r, 'Stripe session'))));
  (list.data || [])
    .filter(s => s.payment_status === 'paid' && s.metadata && s.metadata.pass && !known.has(s.id))
    .forEach(recordPaidSession_);
}

function recordPaidSession_(session) {
  return withLock_(() => {
    const t = membersTable_();
    const existing = t.rows.find(r => String(get_(t, r, 'Stripe session')) === session.id);
    if (existing) return memberSummary_(t, existing, true);

    const m = session.metadata || {};
    const pass = getPasses_()[m.pass] || { key: m.pass, label: m.pass };
    const email = (session.customer_details && session.customer_details.email) || session.customer_email || '';
    const start = m.start || nextStartFor_(t, email, m.pin, m.last);
    const expires = expiryFor_(start, pass.key);
    const memberNo = Number(m.member_no) || existingMemberNo_(t, email, m.pin, m.last) || nextMemberNo_(t);

    const row = newRow_(t, {
      'Member #': memberNo,
      'Added': new Date(),
      'First name': safe_(m.first),
      'Last name': safe_(m.last),
      'Email': safe_(email),
      'Phone': safe_(m.phone),
      'NZHGPA PIN': safe_(m.pin),
      'Type': 'Recreational',
      'Pass': pass.label,
      'Starts': toDate_(start),
      'Expires': toDate_(expires),
      'Show on list': 'Yes',
      'Source': 'Stripe',
      'Stripe session': session.id,
      'Amount paid': (session.amount_total || 0) / 100,
    });
    t.sh.appendRow(row);
    clearMembersCache_();
    maybeEmail_(email, m.first, memberNo, pass.label, start, expires);
    return memberSummary_(t, row, false);
  });
}

/**
 * No start date chosen: start today, or the day after the pilot's current cover ends.
 * Only cover that includes today counts, so a pass pre-booked for later never pushes today's pass back.
 */
function nextStartFor_(t, email, pin, last) {
  const today = todayYmd_();
  const passes = t.rows.filter(r => samePilot_(t, r, email, pin, last)).map(r => passOf_(t, r));
  const end = coverageEnd_(passes, today);
  return end ? addDays_(end, 1) : today;
}

function passOf_(t, r) {
  return { s: ymd_(get_(t, r, 'Starts')), e: ymd_(get_(t, r, 'Expires')), row: r };
}

/** Last day of unbroken cover that includes `day`, chaining passes that follow on with no gap. '' if nothing covers `day`. */
function coverageEnd_(passes, day) {
  let end = '';
  passes.forEach(p => { if (p.e && p.e >= day && (!p.s || p.s <= day) && p.e > end) end = p.e; });
  if (!end) return '';
  for (let grew = true; grew;) {
    grew = false;
    passes.forEach(p => {
      if (p.e > end && p.s && p.s <= addDays_(end, 1)) { end = p.e; grew = true; }
    });
  }
  return end;
}

/* ---------------------------------------------------------------- commercial: company password */

function commercialSignup_(b) {
  const d = validatePilot_(b, { pinRequired: true });
  if (d.error) return d;
  const companyName = clean_(b.company);
  if (!companyName) return fail_('Choose your company.', 'company');

  const cache = CacheService.getScriptCache();
  const failKey = 'fails:' + companyName.toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) return fail_('Too many incorrect passwords. Try again in 15 minutes or contact your manager.', 'password');

  const company = getCompanies_().find(c => c.name.toLowerCase() === companyName.toLowerCase());
  if (!company || !company.active || String(b.password || '').trim() !== company.password) {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS);
    return fail_(PASSWORD_ERROR, 'password');
  }
  cache.remove(failKey);

  return withLock_(() => {
    const t = membersTable_();
    const today = todayYmd_();
    const current = t.rows.find(r => get_(t, r, 'Type') === 'Commercial'
      && cleanPin_(get_(t, r, 'NZHGPA PIN')) === d.pin
      && String(get_(t, r, 'Company')).toLowerCase() === company.name.toLowerCase()
      && ymd_(get_(t, r, 'Expires')) >= today);
    if (current) return memberSummary_(t, current, true);

    const expires = addDays_(today, (company.days || 365) - 1);
    const memberNo = existingMemberNo_(t, d.email, d.pin, d.last) || nextMemberNo_(t);
    const row = newRow_(t, {
      'Member #': memberNo,
      'Added': new Date(),
      'First name': safe_(d.first),
      'Last name': safe_(d.last),
      'Email': safe_(d.email),
      'Phone': safe_(d.phone),
      'NZHGPA PIN': safe_(d.pin),
      'Type': 'Commercial',
      'Pass': 'Commercial',
      'Company': company.name,
      'Starts': toDate_(today),
      'Expires': toDate_(expires),
      'Show on list': 'Yes',
      'Source': 'Commercial form',
    });
    t.sh.appendRow(row);
    clearMembersCache_();
    maybeEmail_(d.email, d.first, memberNo, 'commercial membership (' + company.name + ')', today, expires);
    return memberSummary_(t, row, false);
  });
}

/* ---------------------------------------------------------------- member numbers */

/**
 * The number this pilot already has, or 0. A match needs the same NZHGPA PIN or email, and the
 * same last name, so a mistyped PIN or a shared family email never borrows someone else's number.
 */
function existingMemberNo_(t, email, pin, last) {
  const numbered = t.rows.filter(r => Number(get_(t, r, 'Member #')) && sameLast_(t, r, last));
  const byPin = pin && numbered.find(r => cleanPin_(get_(t, r, 'NZHGPA PIN')) === String(pin));
  if (byPin) return Number(get_(t, byPin, 'Member #'));
  const byEmail = email && numbered.find(r => String(get_(t, r, 'Email')).toLowerCase() === String(email).toLowerCase());
  return byEmail ? Number(get_(t, byEmail, 'Member #')) : 0;
}

/** Last names agree, ignoring case, spaces and punctuation. A blank on either side doesn't block a match. */
function sameLast_(t, r, last) {
  const norm = v => String(v == null ? '' : v).replace(/^'/, '').toLowerCase().replace(/[^a-z\u00c0-\u024f]/g, '');
  const a = norm(get_(t, r, 'Last name')), b = norm(last);
  return !a || !b || a === b;
}

/** Next unused number. Never reuses one, even if a row is deleted or a checkout is abandoned. Call inside the lock. */
function nextMemberNo_(t) {
  const props = PropertiesService.getScriptProperties();
  const highest = t.rows.reduce((max, r) => Math.max(max, Number(get_(t, r, 'Member #')) || 0), 0);
  const n = Math.max(FIRST_MEMBER_NO, highest + 1, Number(props.getProperty('NEXT_MEMBER_NO')) || 0);
  props.setProperty('NEXT_MEMBER_NO', String(n + 1));
  return n;
}

/** Fills blank Member # cells. Returns how many rows were numbered. */
function assignMissingNumbers_(t) {
  const col = t.map['Member #'];
  let count = 0;
  t.rows.forEach(r => {
    if (Number(r[col])) return;
    r[col] = existingMemberNo_(t, get_(t, r, 'Email'), cleanPin_(get_(t, r, 'NZHGPA PIN')), get_(t, r, 'Last name')) || nextMemberNo_(t);
    count++;
  });
  if (count) t.sh.getRange(2, col + 1, t.rows.length, 1).setValues(t.rows.map(r => [r[col]]));
  return count;
}

function samePilot_(t, r, email, pin, last) {
  const sameEmail = email && String(get_(t, r, 'Email')).toLowerCase() === String(email).toLowerCase();
  const samePin = pin && cleanPin_(get_(t, r, 'NZHGPA PIN')) === String(pin);
  return !!(sameEmail || samePin) && sameLast_(t, r, last);
}

function cleanPin_(v) {
  return String(v == null ? '' : v).replace(/^'/, '').trim();
}

/* ---------------------------------------------------------------- public member list */

/**
 * Public list: one entry per pilot.
 *   current    a pass covers today; expiry = end of unbroken cover
 *   prebooked  nothing covers today but a pass starts later; shows the first start date
 *   expired    their last pass ended within SHOW_EXPIRED_DAYS
 */
function currentMembers_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('members');
  if (hit) return JSON.parse(hit);
  const stamp = cache.get('members-stamp');
  const t = membersTable_();
  const today = todayYmd_();
  const oldest = addDays_(today, -SHOW_EXPIRED_DAYS);
  const groups = new Map();
  t.rows.forEach(r => {
    if (String(get_(t, r, 'Show on list')).trim().toLowerCase() === 'no') return;
    const name = (String(get_(t, r, 'First name')).trim() + ' ' + String(get_(t, r, 'Last name')).trim()).replace(/^'/, '').trim();
    const p = passOf_(t, r);
    if (!name || !p.e || p.e < oldest) return;
    const no = Number(get_(t, r, 'Member #')) || null;
    const key = no ? '#' + no : name.toLowerCase();
    if (!groups.has(key)) groups.set(key, { no, name, passes: [] });
    groups.get(key).passes.push(p);
  });
  const list = [...groups.values()].map(g => {
    let status = 'current', starts, expires = coverageEnd_(g.passes, today);
    const future = g.passes.filter(p => p.s && p.s > today);
    if (expires) {
      starts = g.passes.filter(p => (!p.s || p.s <= today) && p.e >= today).map(p => p.s).sort()[0] || '';
    } else if (future.length) {
      status = 'prebooked';
      starts = future.map(p => p.s).sort()[0];
      expires = coverageEnd_(g.passes, starts);
    } else {
      status = 'expired';
      const latest = g.passes.slice().sort((a, b) => (a.e < b.e ? 1 : -1))[0];
      starts = latest.s;
      expires = latest.e;
    }
    const last = g.passes.filter(p => p.e <= expires).sort((a, b) => (a.e < b.e ? 1 : -1))[0] || g.passes[0];
    const pin = g.passes.map(p => cleanPin_(get_(t, p.row, 'NZHGPA PIN'))).filter(Boolean).pop() || '';
    return {
      no: g.no, name: g.name, pin,
      type: String(get_(t, last.row, 'Type') || ''),
      company: String(get_(t, last.row, 'Company') || ''),
      starts, expires, status,
    };
  }).sort((a, b) => a.name.localeCompare(b.name));
  // Only cache this copy if nobody signed up or paid while we were reading the sheet.
  if (cache.get('members-stamp') === stamp) cache.put('members', JSON.stringify(list), 60);
  return list;
}

/** Called after every write: drops the cached list and marks it as changed, so a read that started earlier can't re-cache old data. */
function clearMembersCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove('members');
  cache.put('members-stamp', String(Date.now()) + Math.random(), 21600);
}

/* ---------------------------------------------------------------- sheet helpers (columns found by header name) */

/** Reads the Members tab. Adds any missing columns; a new Member # column goes first and gets numbered. */
function membersTable_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB.members);
  if (!sh) throw Object.assign(new Error('Members tab missing'), { userMessage: 'The member list is not set up yet.' });
  let headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim()) : [];
  let addedNumberColumn = false;
  if (headers.indexOf('Member #') < 0) {
    sh.insertColumnBefore(1);
    sh.getRange(1, 1).setValue('Member #').setFontWeight('bold');
    sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('0');
    headers = ['Member #'].concat(headers);
    addedNumberColumn = true;
  }
  MEMBER_COLS.forEach(name => {
    if (headers.indexOf(name) < 0) {
      sh.getRange(1, headers.length + 1).setValue(name).setFontWeight('bold');
      headers.push(name);
    }
  });
  const map = {};
  headers.forEach((h, i) => { if (h && !(h in map)) map[h] = i; });
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, headers.length).getValues() : [];
  const t = { sh, map, rows, width: headers.length };
  if (addedNumberColumn && rows.length) assignMissingNumbers_(t);
  return t;
}

function get_(t, row, name) {
  return row[t.map[name]];
}

function newRow_(t, values) {
  const row = new Array(t.width).fill('');
  Object.keys(values).forEach(name => { row[t.map[name]] = values[name]; });
  return row;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getCompanies_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB.companies);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => r[0])
    .map(r => ({
      name: String(r[0]).trim(),
      password: String(r[1]).trim(),
      active: String(r[2]).trim().toLowerCase() !== 'no',
      days: Number(r[3]) || 365,
    }));
}

function getPasses_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TAB.passes);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    const key = String(r[0]).trim().toLowerCase();
    if (key && Number(r[1]) > 0) out[key] = { key, price: Number(r[1]), label: String(r[2] || key) };
  });
  return out;
}

function memberSummary_(t, row, already) {
  return {
    ok: true,
    already,
    memberNo: Number(get_(t, row, 'Member #')) || null,
    name: String(get_(t, row, 'First name')).replace(/^'/, ''),
    pass: String(get_(t, row, 'Pass')),
    company: String(get_(t, row, 'Company') || ''),
    starts: ymd_(get_(t, row, 'Starts')),
    expires: ymd_(get_(t, row, 'Expires')),
  };
}

/* ---------------------------------------------------------------- validation */

function validatePilot_(b, opts) {
  const d = {
    first: clean_(b.first), last: clean_(b.last), email: clean_(b.email).toLowerCase(),
    phone: clean_(b.phone), pin: clean_(b.pin).replace(/\s+/g, ''),
  };
  if (!d.first) return fail_('Enter your first name.', 'first');
  if (!d.last) return fail_('Enter your last name.', 'last');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.email)) return fail_('Check your email address.', 'email');
  if (!/^[+\d][\d\s()-]{5,}$/.test(d.phone)) return fail_('Check your phone number.', 'phone');
  if (opts.pinRequired && !d.pin) return fail_('Enter your NZHGPA PIN.', 'pin');
  if (d.pin && !/^\d{1,10}$/.test(d.pin)) return fail_('Your NZHGPA PIN should be numbers only.', 'pin');
  if (!b.agreed) return fail_('Agree to the field rules and the members list to continue.', 'agreed');
  return d;
}

function fail_(error, field) {
  return { ok: false, error, field: field || null };
}

function clean_(v) {
  return String(v == null ? '' : v).trim().slice(0, 200);
}

/** Stops anything a pilot types from being treated as a spreadsheet formula. */
function safe_(v) {
  const s = clean_(v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

/* ---------------------------------------------------------------- dates (yyyy-MM-dd strings in NZ time) */

function todayYmd_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function ymd_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return s;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0') : '';
}

function toDate_(ymd) {
  return Utilities.parseDate(ymd, TZ, 'yyyy-MM-dd');
}

function addDays_(ymd, n) {
  const d = new Date(ymd + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

function addMonths_(ymd, n) {
  const [y, m, day] = ymd.split('-').map(Number);
  const target = new Date(Date.UTC(y, m - 1 + n, 1, 12));
  const lastDay = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0, 12)).getUTCDate();
  target.setUTCDate(Math.min(day, lastDay));
  return target.toISOString().slice(0, 10);
}

function expiryFor_(start, pass) {
  if (pass === 'day') return start;
  if (pass === 'week') return addDays_(start, 6);
  if (pass === 'month') return addDays_(addMonths_(start, 1), -1);
  if (pass === 'year') return addDays_(addMonths_(start, 12), -1);
  return start;
}

/* ---------------------------------------------------------------- Stripe and email */

function stripe_(method, path, params) {
  const key = prop_('STRIPE_SECRET_KEY');
  const options = { method, headers: { Authorization: 'Bearer ' + key }, muteHttpExceptions: true };
  // Apps Script sends numbers as "1.0", which Stripe rejects, so send every value as text.
  if (method === 'post') options.payload = Object.fromEntries(Object.entries(params).map(([k, v]) => [k, String(v)]));
  const res = UrlFetchApp.fetch('https://api.stripe.com' + path, options);
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() >= 300) {
    console.error('Stripe error', res.getResponseCode(), body.error && body.error.message);
    throw Object.assign(new Error('Stripe ' + res.getResponseCode()), { status: res.getResponseCode(), userMessage: 'The payment system did not respond properly. Please try again.' });
  }
  return body;
}

function prop_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  if (!v) throw Object.assign(new Error('Missing script property ' + name), { userMessage: 'Online sign-up is not set up yet.' });
  return v;
}

function maybeEmail_(to, first, memberNo, passLabel, start, expires) {
  if (PropertiesService.getScriptProperties().getProperty('SEND_EMAILS') !== 'yes' || !to) return;
  const nice = s => Utilities.formatDate(toDate_(s), TZ, 'd MMMM yyyy');
  const valid = start === expires ? nice(start) : nice(start) + ' to ' + nice(expires);
  MailApp.sendEmail({
    to,
    replyTo: 'membership@flightpark.co.nz',
    name: 'Flight Park Queenstown',
    subject: "You're on the Flight Park list (member #" + memberNo + ')',
    body: 'Hi ' + (first || 'there') + ',\n\nYour ' + passLabel + ' is confirmed, valid ' + valid +
      '.\nYour member number is ' + memberNo + '.\n\nPlease follow the field rules and fly safe.\n\nFlight Park Queenstown',
  });
}
