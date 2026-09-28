const express = require('express');
const { chromium } = require('playwright');
const http = require('http');
const { Server } = require('socket.io');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');

const app = express();
const server = http.createServer(app);
const io = new Server(server);

const OLLAMA_URL = process.env.OLLAMA_URL || 'http://localhost:11434';
let OLLAMA_MODEL = process.env.OLLAMA_MODEL || 'llama3';

const upload = multer({ dest: path.join(__dirname, 'uploads') });
if (!fs.existsSync(path.join(__dirname, 'uploads'))) fs.mkdirSync(path.join(__dirname, 'uploads'));

let cachedProfile = null;
let activeBrowser = null;
let activePage = null;
let agentRunning = false;
let humanConfirmResolve = null;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ══════════════════════════════════════════
// EXCEL — UPLOAD & SHEET UPDATE
// ══════════════════════════════════════════

app.post('/api/upload-profile', upload.single('excel'), (req, res) => {
  try {
    if (!req.file) return res.json({ error: 'No file uploaded' });
    const wb = XLSX.readFile(req.file.path);
    const profileSheet = wb.Sheets['Company Profile'];
    if (!profileSheet) return res.json({ error: 'Sheet "Company Profile" not found' });

    // row 1=banner, row 2=headers, row 3+=data → range:2 skips first 2 rows (0-indexed)
    const rows = XLSX.utils.sheet_to_json(profileSheet, { header: ['field','value','hint'], range: 2, defval: '' });
    const lines = [];
    for (const r of rows) {
      const f = (r.field||'').toString().trim();
      const v = (r.value||'').toString().trim();
      if (f && v && f !== '__SECTION__') lines.push(`${f}: ${v}`);
    }
    const profileText = lines.join('\n');

    const grantSheet = wb.Sheets['Grant Tracker'];
    let grants = [];
    if (grantSheet) {
      const gRows = XLSX.utils.sheet_to_json(grantSheet, {
        header: ['row_id','grant_name','url','deadline','amount','status','filled_date','filled_count','notes'],
        range: 2, defval: ''
      });
      grants = gRows.filter(r => r.url && r.grant_name && (r.status||'').toString().toLowerCase().trim() === 'pending');
    }

    const savedPath = path.join(__dirname, 'uploads', 'profile.xlsx');
    fs.copyFileSync(req.file.path, savedPath);
    try { fs.unlinkSync(req.file.path); } catch {}
    cachedProfile = { profileText, grants, filePath: savedPath };
    res.json({ success: true, fieldCount: lines.length, pendingGrants: grants.length, grants: grants.map(g => ({ name: g.grant_name, url: g.url })) });
  } catch (e) { res.json({ error: e.message }); }
});

app.get('/api/profile', (req, res) => {
  if (!cachedProfile) return res.json({ loaded: false });
  res.json({ loaded: true, fieldCount: cachedProfile.profileText.split('\n').length, pendingGrants: cachedProfile.grants.length, grants: cachedProfile.grants.map(g => ({ name: g.grant_name, url: g.url })) });
});

// Write back to Grant Tracker — columns: A=row_id B=grant_name C=url D=deadline E=amount F=status G=filled_date H=filled_count I=notes
function updateSheet(grantName, status, filledCount, total, notes) {
  if (!cachedProfile?.filePath) return;
  try {
    const wb = XLSX.readFile(cachedProfile.filePath);
    const ws = wb.Sheets['Grant Tracker'];
    if (!ws) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 2; R <= range.e.r; R++) { // data starts at row index 2 (row 3 in Excel)
      const nameCell = ws[XLSX.utils.encode_cell({ r: R, c: 1 })]; // B = grant_name
      if (!nameCell) continue;
      if ((nameCell.v||'').toString().trim().toLowerCase() === grantName.trim().toLowerCase()) {
        ws[XLSX.utils.encode_cell({ r: R, c: 5 })] = { v: status, t: 's' };
        ws[XLSX.utils.encode_cell({ r: R, c: 6 })] = { v: new Date().toLocaleDateString('en-IN'), t: 's' };
        ws[XLSX.utils.encode_cell({ r: R, c: 7 })] = { v: `${filledCount}/${total}`, t: 's' };
        ws[XLSX.utils.encode_cell({ r: R, c: 8 })] = { v: notes, t: 's' };
        XLSX.writeFile(wb, cachedProfile.filePath);
        log(`Sheet updated: "${grantName}" → ${status} (${filledCount}/${total})`, 'success');
        io.emit('sheet_updated', { grantName, status, filledCount, total });
        return;
      }
    }
    log(`"${grantName}" not found in sheet`, 'warn');
  } catch (e) { log(`Sheet update failed: ${e.message}`, 'warn'); }
}

// ══════════════════════════════════════════
// OLLAMA
// ══════════════════════════════════════════

async function ollamaChat(prompt) {
  const res = await fetch(`${OLLAMA_URL}/api/chat`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: OLLAMA_MODEL, stream: false, messages: [{ role: 'user', content: prompt }] })
  });
  if (!res.ok) throw new Error(`Ollama not reachable — run: ollama serve`);
  const data = await res.json();
  return data.message?.content?.trim() || '';
}

function log(msg, type = 'info') {
  console.log(`[${type}] ${msg}`);
  io.emit('log', { msg, type });
}

function waitForHuman(label) {
  return new Promise(resolve => {
    humanConfirmResolve = resolve;
    io.emit('human_review', { section: label });
    log(`Paused — review "${label}" then click Continue.`, 'human');
  });
}

io.on('connection', socket => {
  socket.on('human_continue', () => {
    if (humanConfirmResolve) { humanConfirmResolve('continue'); humanConfirmResolve = null; io.emit('human_review_done'); }
  });
  socket.on('stop_agent', () => {
    agentRunning = false;
    if (humanConfirmResolve) { humanConfirmResolve('stop'); humanConfirmResolve = null; }
    log('Stopped.', 'warn');
    io.emit('agent_stopped');
  });
});

// ══════════════════════════════════════════
// UNIVERSAL FIELD SCANNER
// Works for Google Forms, Accubate, standard HTML, React forms
// Bug fixes: handles [role=radio/checkbox], GForms divs, aria-labelledby, strips * from labels
// ══════════════════════════════════════════

async function scanFields(page) {
  await page.waitForTimeout(1500);

  // Wait for any visible interactive element — broad selector catches both native + ARIA
  try {
    await page.waitForSelector(
      'input:not([type=hidden]):not([type=submit]):not([type=button]):not([type=reset]), textarea, select, [role=radio], [role=checkbox], [role=listbox], [role=combobox]',
      { timeout: 10000, state: 'visible' }
    );
  } catch { return []; }

  return await page.evaluate(() => {
    const SKIP_PATTERNS = ['recaptcha','captcha','csrf','_token','honeypot','g-recaptcha','__','nonce','utm_','fbclid','gclid'];

    function shouldSkip(el) {
      const s = [(el.id||''),(el.name||''),(el.className||'')].join(' ').toLowerCase();
      return SKIP_PATTERNS.some(p => s.includes(p));
    }

    function isVisible(el) {
      if (el.type === 'hidden') return false;
      const r = el.getBoundingClientRect();
      const st = window.getComputedStyle(el);
      return r.width > 0 && r.height > 0 &&
        st.display !== 'none' && st.visibility !== 'hidden' &&
        st.opacity !== '0' && !el.closest('[hidden]');
    }

    // FIX: Strip asterisk (*) and extra whitespace from labels
    function cleanLabel(txt) {
      return (txt || '').replace(/\s*\*\s*$/, '').replace(/\s+/g, ' ').trim();
    }

    // FIX: Universal label finder — works for standard HTML AND Google Forms
    function getLabel(el) {
      // 1. aria-labelledby (GForms uses this)
      const labelledBy = el.getAttribute('aria-labelledby');
      if (labelledBy) {
        const parts = labelledBy.split(' ').map(id => document.getElementById(id)?.innerText?.trim()).filter(Boolean);
        if (parts.length) return cleanLabel(parts.join(' '));
      }
      // 2. aria-label
      if (el.getAttribute('aria-label')) return cleanLabel(el.getAttribute('aria-label'));
      // 3. Explicit <label for="id">
      if (el.id) {
        const lbl = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
        if (lbl) return cleanLabel(lbl.innerText);
      }
      // 4. Wrapping label
      const wrapLabel = el.closest('label');
      if (wrapLabel) return cleanLabel(wrapLabel.innerText.replace(el.value||'',''));
      // 5. GForms: question heading is in .M7eMe or .HoXoMd inside [data-params] container
      const gformsQ = el.closest('[jsmodel], [data-params], .Qr7Oae, .geS5n');
      if (gformsQ) {
        const heading = gformsQ.querySelector('.M7eMe, .HoXoMd, .z65CN, [role=heading]');
        if (heading) return cleanLabel(heading.innerText);
      }
      // 6. Walk up DOM — look for preceding heading/label
      let node = el.parentElement;
      for (let i = 0; i < 8; i++) {
        if (!node) break;
        const cands = node.querySelectorAll('label, legend, [role=heading], [class*="label"], [class*="title"], [class*="question"]');
        for (const c of cands) {
          const txt = c.innerText?.trim();
          if (txt && txt.length > 1 && txt.length < 200 && !c.contains(el)) return cleanLabel(txt);
        }
        node = node.parentElement;
      }
      // 7. Placeholder / name fallback
      if (el.placeholder?.length > 1) return el.placeholder;
      if (el.name) return el.name.replace(/[-_[\]]/g, ' ').trim();
      return 'unnamed field';
    }

    // FIX: Stable selector that doesn't break when conditional fields appear
    function getSelector(el) {
      if (el.id) return `#${CSS.escape(el.id)}`;
      if (el.name) return `[name="${CSS.escape(el.name)}"]`;
      // Build a path using data-params or jsname for GForms stability
      if (el.getAttribute('data-params')) return `[data-params="${el.getAttribute('data-params').slice(0,40)}"]`;
      if (el.getAttribute('jsname')) return `[jsname="${el.getAttribute('jsname')}"]`;
      // Last resort: stable ancestor + role
      const role = el.getAttribute('role');
      if (role) {
        const tag = el.tagName.toLowerCase();
        const all = Array.from(document.querySelectorAll(`${tag}[role="${role}"]`));
        const idx = all.indexOf(el);
        if (idx >= 0) return `${tag}[role="${role}"]:nth-of-type(${idx+1})`;
      }
      const tag = el.tagName.toLowerCase();
      const all = Array.from(document.querySelectorAll(tag));
      return `${tag}:nth-of-type(${all.indexOf(el)+1})`;
    }

    function inferPurpose(el, label) {
      const sig = [label, el.name||'', el.id||'', el.placeholder||'', el.getAttribute('autocomplete')||'', el.getAttribute('inputmode')||''].join(' ').toLowerCase();
      if (/\bemail\b/.test(sig)) return 'email';
      if (/phone|mobile|tel|contact.?no/.test(sig)) return 'phone';
      if (/website|url|http|www/.test(sig)) return 'website';
      if (/\bpan\b/.test(sig)) return 'pan';
      if (/\bgst\b/.test(sig)) return 'gst';
      if (/pin.?code|postal|zip/.test(sig)) return 'pincode';
      if (/\bstate\b/.test(sig)) return 'state';
      if (/\bcity\b/.test(sig)) return 'city';
      if (/\bcountry\b/.test(sig)) return 'country';
      if (/\baddress\b/.test(sig)) return 'address';
      if (/founded|incorporat/.test(sig)) return 'founding';
      if (/revenue|turnover|sales/.test(sig)) return 'revenue';
      if (/team.?size|employee|staff/.test(sig)) return 'team_size';
      if (/\bname\b/.test(sig) && /company|org|enterprise|startup/.test(sig)) return 'company_name';
      if (/\bname\b/.test(sig) && /applicant|founder|person|contact/.test(sig)) return 'founder_name';
      if (/position|designation|title|role/.test(sig)) return 'designation';
      if (/descri|mission|about|overview|summary/.test(sig)) return 'description';
      if (/impact|beneficiar|outcome/.test(sig)) return 'impact';
      if (/funding|investment|raised|capital/.test(sig)) return 'funding';
      if (/grant.?use|utiliz|spend|allocat/.test(sig)) return 'grant_use';
      if (/market|sector|industry/.test(sig)) return 'market';
      if (/regist|cin|llpin/.test(sig)) return 'registration';
      if (/\byear\b/.test(sig)) return 'year';
      return 'text';
    }

    const results = [];
    const seen = new Set();

    // ── NATIVE HTML inputs ──
    document.querySelectorAll(
      'input:not([type=hidden]):not([type=submit]):not([type=button]):not([type=reset]):not([type=image]), textarea, select'
    ).forEach(el => {
      if (!isVisible(el) || shouldSkip(el)) return;
      // Skip terms checkboxes
      if (el.type === 'checkbox') {
        const lbl = getLabel(el).toLowerCase();
        if (['agree','accept','terms','consent','privacy','policy'].some(k => lbl.includes(k))) return;
      }
      // Deduplicate radio groups by name
      if (el.type === 'radio' && el.name) {
        if (seen.has('radio:' + el.name)) return;
        seen.add('radio:' + el.name);
      }

      const type = el.tagName === 'SELECT' ? 'select' : el.tagName === 'TEXTAREA' ? 'textarea' : (el.type||'text');
      const label = getLabel(el);
      const sel = getSelector(el);
      if (seen.has(sel)) return;
      seen.add(sel);

      const maxLength = (el.maxLength > 0 && el.maxLength < 524288) ? el.maxLength : null;

      let options = [];
      if (el.tagName === 'SELECT') {
        options = Array.from(el.options).map(o => o.text.trim()).filter(t => t && !['','--','select','choose','please select'].includes(t.toLowerCase()));
      }
      let radioOptions = [];
      if (el.type === 'radio' && el.name) {
        document.querySelectorAll(`input[name="${CSS.escape(el.name)}"]`).forEach(r => {
          const lbl = document.querySelector(`label[for="${CSS.escape(r.id)}"]`);
          radioOptions.push(cleanLabel(lbl?.innerText || r.value));
        });
      }

      results.push({ selector: sel, label, purpose: inferPurpose(el, label), type, isFile: el.type==='file', isGforms: false, options, radioOptions, required: el.required, name: el.name||'', maxLength, currentValue: el.value?.trim()||'', placeholder: el.placeholder||'' });
    });

    // ── GOOGLE FORMS ARIA elements ──
    // GForms uses [role=radio], [role=checkbox], [role=listbox] inside question containers
    // Each question container has class .Qr7Oae or [data-params]
    const gformsQuestions = document.querySelectorAll('.Qr7Oae, [jsmodel][data-params]');
    gformsQuestions.forEach(qContainer => {
      // Get the question title
      const titleEl = qContainer.querySelector('.M7eMe, .HoXoMd, .z65CN, [role=heading]');
      const label = cleanLabel(titleEl?.innerText || '');
      if (!label) return;

      const required = !!qContainer.querySelector('[aria-required="true"], .vnumgf');

      // Text / Long text input
      const textInput = qContainer.querySelector('input[type=text]:not([type=hidden]), textarea');
      if (textInput && isVisible(textInput) && !seen.has(getSelector(textInput))) {
        const sel = getSelector(textInput);
        seen.add(sel);
        const maxLength = (textInput.maxLength > 0 && textInput.maxLength < 524288) ? textInput.maxLength : null;
        const type = textInput.tagName === 'TEXTAREA' ? 'textarea' : 'text';
        results.push({ selector: sel, label, purpose: inferPurpose(textInput, label), type, isFile: false, isGforms: true, options: [], radioOptions: [], required, name: textInput.name||'', maxLength, currentValue: textInput.value?.trim()||'', placeholder: textInput.placeholder||'' });
        return;
      }

      // Radio group (multiple choice)
      const radioItems = qContainer.querySelectorAll('[role=radio]');
      if (radioItems.length > 0) {
        const groupKey = 'gf-radio-' + label.slice(0,30);
        if (!seen.has(groupKey)) {
          seen.add(groupKey);
          const radioOptions = Array.from(radioItems).map(r => {
            return cleanLabel(r.querySelector('.aDTYNe, .Y6Myld, span')?.innerText || r.getAttribute('data-value') || r.innerText);
          }).filter(Boolean);
          const hasOther = !!qContainer.querySelector('[data-other-option], [jsname*="other"]');
          if (hasOther) radioOptions.push('Other');
          // Use the first radio item's selector — we'll click by label text at fill time
          results.push({ selector: `[role=radio]`, label, purpose: inferPurpose(radioItems[0], label), type: 'gforms-radio', isFile: false, isGforms: true, options: [], radioOptions, required, name: groupKey, maxLength: null, currentValue: '', placeholder: '', gformsContainer: label });
        }
        return;
      }

      // Checkbox group (multiple answers)
      const checkItems = qContainer.querySelectorAll('[role=checkbox]');
      if (checkItems.length > 0) {
        const groupKey = 'gf-check-' + label.slice(0,30);
        if (!seen.has(groupKey)) {
          seen.add(groupKey);
          const checkOptions = Array.from(checkItems).map(c => {
            return cleanLabel(c.querySelector('.aDTYNe, .Y6Myld, span')?.innerText || c.getAttribute('data-value') || c.innerText);
          }).filter(Boolean);
          results.push({ selector: `[role=checkbox]`, label, purpose: inferPurpose(checkItems[0], label), type: 'gforms-checkbox', isFile: false, isGforms: true, options: [], radioOptions: checkOptions, required, name: groupKey, maxLength: null, currentValue: '', placeholder: '', gformsContainer: label });
        }
        return;
      }

      // Dropdown (GForms uses [role=listbox] or [role=option])
      const dropdown = qContainer.querySelector('[role=listbox], select');
      if (dropdown && isVisible(dropdown) && !seen.has(getSelector(dropdown))) {
        const sel = getSelector(dropdown);
        seen.add(sel);
        const opts = Array.from(qContainer.querySelectorAll('[role=option], option')).map(o => cleanLabel(o.innerText)).filter(t => t && t !== 'Choose');
        results.push({ selector: sel, label, purpose: inferPurpose(dropdown, label), type: 'gforms-select', isFile: false, isGforms: true, options: opts, radioOptions: [], required, name: sel, maxLength: null, currentValue: '', placeholder: '', gformsContainer: label });
      }
    });

    return results;
  });
}

// ══════════════════════════════════════════
// SMART ANSWER GENERATION
// ══════════════════════════════════════════

async function generateAnswers(fields, companyProfile, grantName, sectionName, sectionNum, context) {
  const fieldList = fields.map((f, i) => {
    if (f.isFile) return `${i+1}. [FILE] "${f.label}" → SKIP`;
    let line = `${i+1}. [${f.type.toUpperCase()}] "${f.label}"`;
    line += `\n   Purpose hint: ${f.purpose}`;
    if (f.placeholder) line += `\n   Placeholder: "${f.placeholder}"`;
    if (f.maxLength) line += `\n   MAX ${f.maxLength} characters — strictly do not exceed`;
    if (f.options?.length) line += `\n   OPTIONS (copy exactly): ${f.options.map(o=>`"${o}"`).join(' | ')}`;
    if (f.radioOptions?.length) line += `\n   CHOICES (copy exactly): ${f.radioOptions.map(o=>`"${o}"`).join(' | ')}`;
    if (f.required) line += `\n   REQUIRED`;
    if (f.currentValue) line += `\n   Already has: "${f.currentValue}"`;
    return line;
  }).join('\n\n');

  const prompt = `You are filling a grant application form. Section ${sectionNum}: "${sectionName}" of "${grantName}".

COMPANY PROFILE:
${companyProfile}

${context ? `ALREADY ANSWERED:\n${context}\n` : ''}
FIELDS TO FILL:
${fieldList}

RULES:
- text/textarea: write a real answer using profile data. Stay under max characters if given.
- SELECT / gforms-select: value MUST exactly match one listed OPTION (copy text exactly)
- radio / gforms-radio: value MUST exactly match one listed CHOICE (copy text exactly)
- gforms-checkbox: list the choices to select, comma-separated, e.g. "Technology, Education"
- checkbox: "true" or "false"
- file: "SKIP"
- email: exact email from profile
- phone: digits only from profile
- website: full URL with https://
- year: 4-digit number only
- number/pincode: digits only
- If already has a correct value, repeat it unchanged

Return ONLY a JSON array: [{"index":0,"value":"answer"},...]
No markdown. No explanation. Raw JSON only.`;

  try {
    log('Generating answers...', 'ai');
    const raw = await ollamaChat(prompt);
    const match = raw.replace(/```json|```/g,'').trim().match(/\[[\s\S]*\]/);
    if (!match) throw new Error('No JSON in AI response');
    const answers = JSON.parse(match[0]);
    return fields.map((f, i) => ({ ...f, answer: answers.find(a => Number(a.index) === i)?.value || '' }));
  } catch(e) {
    log('AI error: ' + e.message, 'error');
    return fields.map(f => ({ ...f, answer: '' }));
  }
}

// ══════════════════════════════════════════
// UNIVERSAL FIELD FILLER
// FIX: Handles GForms ARIA elements, "Other" option, validates values
// ══════════════════════════════════════════

async function fillField(page, field) {
  if (field.isFile) return { status: 'skipped', reason: 'file upload' };
  if (!field.answer || field.answer.toUpperCase() === 'SKIP') return { status: 'skipped', reason: 'no answer' };

  let answer = field.answer.toString().trim();

  // Enforce constraints
  if (field.maxLength && answer.length > field.maxLength) answer = answer.slice(0, field.maxLength);
  if (field.purpose === 'phone' || field.purpose === 'pincode') answer = answer.replace(/\D/g, '');
  if (field.purpose === 'year') { const m = answer.match(/\d{4}/); if (m) answer = m[0]; }

  try {
    // ── GOOGLE FORMS RADIO ──
    if (field.type === 'gforms-radio') {
      const clicked = await page.evaluate((labelText, answerText) => {
        const lower = answerText.toLowerCase().trim();
        // Find the question container by its title
        const containers = document.querySelectorAll('.Qr7Oae, [jsmodel][data-params]');
        for (const container of containers) {
          const titleEl = container.querySelector('.M7eMe, .HoXoMd, .z65CN, [role=heading]');
          const title = (titleEl?.innerText||'').replace(/\s*\*\s*$/, '').trim();
          if (title.toLowerCase() !== labelText.toLowerCase()) continue;
          // Found the right question — now click the matching radio
          const radios = container.querySelectorAll('[role=radio]');
          for (const r of radios) {
            const optText = (r.querySelector('.aDTYNe, .Y6Myld, span')?.innerText || r.getAttribute('data-value') || r.innerText || '').trim().toLowerCase();
            if (optText === lower || optText.includes(lower) || lower.includes(optText)) {
              r.click();
              return optText;
            }
          }
          // Fallback: first option
          if (radios[0]) { radios[0].click(); return radios[0].innerText?.trim() + ' (first)'; }
        }
        return null;
      }, field.label, answer);

      if (!clicked) return { status: 'failed', reason: `no radio option matched "${answer}"` };
      await page.waitForTimeout(300);
      // Check if "Other" was selected — if so, type in the text box
      if (clicked.toLowerCase().includes('other')) {
        const otherInput = await page.$(`[data-other-option] input, [jsname*="other"] input, input[placeholder*="other" i]`);
        if (otherInput) { await otherInput.click(); await otherInput.fill(answer); }
      }
      return { status: 'filled', value: clicked };
    }

    // ── GOOGLE FORMS CHECKBOX ──
    if (field.type === 'gforms-checkbox') {
      const chosen = answer.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
      const clicked = await page.evaluate((labelText, chosenList) => {
        const containers = document.querySelectorAll('.Qr7Oae, [jsmodel][data-params]');
        for (const container of containers) {
          const titleEl = container.querySelector('.M7eMe, .HoXoMd, .z65CN, [role=heading]');
          const title = (titleEl?.innerText||'').replace(/\s*\*\s*$/, '').trim();
          if (title.toLowerCase() !== labelText.toLowerCase()) continue;
          const items = container.querySelectorAll('[role=checkbox]');
          const ticked = [];
          items.forEach(item => {
            const optText = (item.querySelector('.aDTYNe, .Y6Myld, span')?.innerText || item.innerText || '').trim().toLowerCase();
            const shouldCheck = chosenList.some(c => optText.includes(c) || c.includes(optText));
            if (shouldCheck) { item.click(); ticked.push(optText); }
          });
          return ticked;
        }
        return [];
      }, field.label, chosen);
      await page.waitForTimeout(300);
      return clicked.length > 0 ? { status: 'filled', value: clicked.join(', ') } : { status: 'failed', reason: 'no checkbox options matched' };
    }

    // ── GOOGLE FORMS SELECT/DROPDOWN ──
    if (field.type === 'gforms-select') {
      const clicked = await page.evaluate((labelText, answerText) => {
        const lower = answerText.toLowerCase().trim();
        const containers = document.querySelectorAll('.Qr7Oae, [jsmodel][data-params]');
        for (const container of containers) {
          const titleEl = container.querySelector('.M7eMe, .HoXoMd, .z65CN, [role=heading]');
          const title = (titleEl?.innerText||'').replace(/\s*\*\s*$/, '').trim();
          if (title.toLowerCase() !== labelText.toLowerCase()) continue;
          // Click the dropdown to open it
          const trigger = container.querySelector('[role=listbox], .DzsxAb, [jsname="VIuuj"]');
          if (trigger) { trigger.click(); return 'opened'; }
        }
        return null;
      }, field.label, answer);

      if (clicked === 'opened') {
        await page.waitForTimeout(500);
        // Now click the matching option
        const selected = await page.evaluate((answerText) => {
          const lower = answerText.toLowerCase().trim();
          const options = document.querySelectorAll('[role=option], .MocG8c [data-value]');
          for (const opt of options) {
            const txt = (opt.innerText || opt.getAttribute('data-value') || '').trim().toLowerCase();
            if (txt === lower || txt.includes(lower) || lower.includes(txt)) { opt.click(); return txt; }
          }
          if (options[0]) { options[0].click(); return options[0].innerText?.trim() + ' (first)'; }
          return null;
        }, answer);
        await page.waitForTimeout(300);
        return selected ? { status: 'filled', value: selected } : { status: 'failed', reason: 'dropdown option not found' };
      }
      // Fallback to native select
    }

    // ── NATIVE SELECT ──
    if (field.type === 'select') {
      let el;
      try { await page.waitForSelector(field.selector, { state: 'visible', timeout: 5000 }); el = await page.$(field.selector); } catch {}
      if (!el) return { status: 'not_found', reason: 'select element not found' };
      await el.evaluate(e => e.scrollIntoView({ block: 'center' }));
      const picked = await el.evaluate((sel, ans) => {
        const lower = ans.toLowerCase().trim();
        for (const opt of sel.options) {
          const t = opt.text.trim().toLowerCase();
          if (t === lower || t.includes(lower) || lower.includes(t)) {
            sel.value = opt.value;
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            return opt.text.trim();
          }
        }
        return null;
      }, answer);
      if (!picked && field.options?.[0]) { await el.selectOption({ label: field.options[0] }); return { status: 'filled', value: field.options[0] + ' (fallback)' }; }
      return picked ? { status: 'filled', value: picked } : { status: 'failed', reason: `no option matched "${answer}"` };
    }

    // ── NATIVE CHECKBOX ──
    if (field.type === 'checkbox') {
      let el;
      try { await page.waitForSelector(field.selector, { state: 'visible', timeout: 5000 }); el = await page.$(field.selector); } catch {}
      if (!el) return { status: 'not_found', reason: 'checkbox not found' };
      const shouldCheck = ['true','yes','1','check','checked','agree'].includes(answer.toLowerCase());
      if (shouldCheck !== await el.isChecked()) await el.click();
      return { status: 'filled', value: shouldCheck ? 'checked' : 'unchecked' };
    }

    // ── NATIVE RADIO ──
    if (field.type === 'radio') {
      const result = await page.evaluate((name, ans) => {
        const lower = ans.toLowerCase().trim();
        const radios = document.querySelectorAll(`input[name="${name}"]`);
        for (const r of radios) {
          const lbl = document.querySelector(`label[for="${r.id}"]`);
          const txt = (lbl?.innerText || r.value || '').toLowerCase().replace(/\s*\*\s*$/, '').trim();
          if (txt === lower || txt.includes(lower) || lower.includes(txt)) { r.click(); return lbl?.innerText?.trim() || r.value; }
        }
        if (radios[0]) { radios[0].click(); return radios[0].value + ' (first)'; }
        return null;
      }, field.name, answer);
      return result ? { status: 'filled', value: result } : { status: 'failed', reason: 'no radio match' };
    }

    // ── TEXT / TEXTAREA ──
    let el;
    try { await page.waitForSelector(field.selector, { state: 'visible', timeout: 5000 }); el = await page.$(field.selector); } catch {
      el = await page.$(`[placeholder="${field.placeholder}"]`) || await page.$(`[aria-label="${field.label}"]`);
    }
    if (!el) return { status: 'not_found', reason: 'text input not found' };

    await el.evaluate(e => e.scrollIntoView({ block: 'center' }));
    await page.waitForTimeout(200);
    await el.click({ clickCount: 3 });
    await page.waitForTimeout(100);
    await el.fill(answer);
    await el.evaluate(e => {
      e.dispatchEvent(new Event('input', { bubbles: true }));
      e.dispatchEvent(new Event('change', { bubbles: true }));
      e.dispatchEvent(new Event('blur', { bubbles: true }));
    });
    const actual = await el.evaluate(e => e.value);
    return { status: 'filled', value: actual || answer };

  } catch(e) {
    return { status: 'error', reason: e.message.slice(0, 120) };
  }
}

// ══════════════════════════════════════════
// TERMS CHECKBOX
// ══════════════════════════════════════════

async function handleTerms(page) {
  const count = await page.evaluate(() => {
    const TERMS = ['agree','accept','terms','condition','consent','privacy','policy','confirm','declare'];
    let n = 0;
    document.querySelectorAll('input[type=checkbox]').forEach(cb => {
      let txt = '';
      if (cb.id) { const l = document.querySelector(`label[for="${CSS.escape(cb.id)}"]`); if (l) txt = l.innerText.toLowerCase(); }
      if (!txt) txt = (cb.closest('label')?.innerText || cb.parentElement?.innerText || '').toLowerCase();
      if (TERMS.some(p => txt.includes(p)) && !cb.checked) { cb.click(); n++; }
    });
    return n;
  });
  if (count > 0) { log(`Auto-checked ${count} terms checkbox(es)`, 'success'); io.emit('terms_checked', { count }); }
  return count;
}

// ══════════════════════════════════════════
// PAGE STATE & NAVIGATION
// FIX: GForms uses submit for Next, catches "Back" too, reads GForms title
// ══════════════════════════════════════════

async function detectPageState(page) {
  return await page.evaluate(() => {
    const NEXT_KW   = ['next','continue','proceed',"let's get started",'lets get started','get started','save & next','save and next','next section'];
    const SUBMIT_KW = ['submit','apply now','apply','finish','complete','send application','submit form'];
    const BACK_KW   = ['back','previous','prev'];
    const allBtns   = Array.from(document.querySelectorAll('button, input[type=submit], [role=button], [jsname]'));
    const matchBtn  = (b, kws) => kws.some(k => (b.innerText||b.value||b.getAttribute('aria-label')||'').toLowerCase().trim().includes(k));
    const nextBtn   = allBtns.find(b => matchBtn(b, NEXT_KW));
    const submitBtn = allBtns.find(b => matchBtn(b, SUBMIT_KW) && !matchBtn(b, NEXT_KW));

    // FIX: GForms title
    let heading = '';
    const gTitle = document.querySelector('.freebirdFormviewerViewHeaderTitle, .freebirdFormviewerViewHeaderTitlePart, [role=heading]');
    if (gTitle?.innerText?.trim()) heading = gTitle.innerText.trim();
    if (!heading) {
      for (const tag of ['h1','h2','h3']) {
        const el = document.querySelector(tag);
        if (el?.innerText?.trim()) { heading = el.innerText.trim(); break; }
      }
    }

    const fields = Array.from(document.querySelectorAll(
      'input:not([type=hidden]):not([type=submit]):not([type=button]):not([type=reset]), textarea, select, [role=radio], [role=checkbox], [role=listbox]'
    )).filter(el => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; }).length;

    // FIX: Detect GForms page progress (e.g. "Page 2 of 4")
    const progress = document.querySelector('.freebirdFormviewerViewNumberedPageProgressBar, [role=progressbar]');
    const progressText = progress?.innerText?.trim() || progress?.getAttribute('aria-label') || '';

    const isConfirmation = document.querySelector('.freebirdFormviewerViewResponseConfirmationMessage, .freebirdFormviewerViewResponseLinkedResponseMessage');

    return {
      heading: (heading || document.title || 'Section').slice(0, 80),
      hasNextBtn: !!nextBtn,
      nextBtnText: (nextBtn?.innerText || nextBtn?.value || '').trim(),
      hasSubmitBtn: !!submitBtn,
      submitBtnText: (submitBtn?.innerText || submitBtn?.value || '').trim(),
      fieldCount: fields,
      url: window.location.href,
      progressText,
      isConfirmation: !!isConfirmation
    };
  });
}

async function clickProceedButton(page) {
  return await page.evaluate(() => {
    const KWS = ['next','continue','proceed',"let's get started",'get started','save & next','save and next','submit','apply'];
    for (const kw of KWS) {
      for (const btn of document.querySelectorAll('button, input[type=submit], [role=button], [jsname]')) {
        const txt = (btn.innerText||btn.value||btn.getAttribute('aria-label')||'').toLowerCase().trim();
        if (txt.includes(kw) && !txt.includes('back') && !txt.includes('previous')) {
          btn.click();
          return txt;
        }
      }
    }
    return null;
  });
}

async function waitForChange(page, prevHeading, prevUrl, ms = 8000) {
  const deadline = Date.now() + ms;
  while (Date.now() < deadline) {
    await page.waitForTimeout(600);
    const s = await detectPageState(page);
    if (s.url !== prevUrl || s.heading !== prevHeading || s.isConfirmation) return s;
  }
  return null;
}

// ══════════════════════════════════════════
// MAIN AGENT LOOP
// ══════════════════════════════════════════

app.post('/api/run', async (req, res) => {
  if (agentRunning) return res.json({ error: 'Agent already running' });
  const { url, grantName, model } = req.body;
  const companyProfile = req.body.companyProfile || cachedProfile?.profileText;
  if (!url) return res.json({ error: 'Missing grant URL' });
  if (!companyProfile) return res.json({ error: 'No company profile — upload your Excel file or type profile manually' });
  if (model) OLLAMA_MODEL = model;

  const resolvedName = grantName || 'Grant Application';
  agentRunning = true;
  res.json({ started: true });

  let totalFilled = 0, totalSkipped = 0, totalFailed = 0, totalFields = 0;

  try {
    log('Checking Ollama...', 'ai');
    await ollamaChat('say: ready').catch(() => { throw new Error('Ollama not running — run: ollama serve'); });
    log(`Ollama OK — model: ${OLLAMA_MODEL}`, 'success');

    log('Launching Chrome...', 'system');
    activeBrowser = await chromium.launch({ headless: false, slowMo: 60, args: ['--start-maximized'] });
    const ctx = await activeBrowser.newContext({ viewport: null });
    activePage = await ctx.newPage();
    activePage.on('pageerror', () => {});

    log(`Opening: ${url}`, 'system');
    io.emit('status', 'navigating');

    // FIX: Use domcontentloaded instead of networkidle — GForms never reaches networkidle
    await activePage.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await activePage.waitForTimeout(3000); // Extra wait for JS-rendered content

    let sectionNum = 0;
    let filledContext = '';
    let knownSelectors = new Set();

    while (agentRunning) {
      sectionNum++;
      await activePage.waitForTimeout(1200);

      const state = await detectPageState(activePage);
      if (state.isConfirmation) {
        log('Form submitted successfully! Confirmation page detected.', 'success');
        io.emit('form_submitted');
        updateSheet(resolvedName, 'submitted', totalFilled, totalFields, `${totalFilled}/${totalFields} fields — submitted successfully`);
        break;
      }

      const sectionLabel = state.heading || `Section ${sectionNum}`;
      const progress = state.progressText ? ` (${state.progressText})` : '';
      log(`── Section ${sectionNum}: "${sectionLabel}"${progress} — ${state.fieldCount} fields ──`, 'system');
      io.emit('section_start', { num: sectionNum, heading: sectionLabel + progress, fieldCount: state.fieldCount });

      // Auto-check terms
      await handleTerms(activePage);

      // Scan fields
      io.emit('status', 'analysing');
      const fields = await scanFields(activePage);
      // Only fill fields not already filled in a previous scan (avoids re-filling on section revisit)
      const newFields = fields.filter(f => !knownSelectors.has(f.selector));

      if (newFields.length === 0 && fields.length === 0) {
        log('No fields found on this section.', 'warn');
        io.emit('section_no_fields', { num: sectionNum, heading: sectionLabel });
      } else {
        const toFill = newFields.length > 0 ? newFields : fields;
        log(`Found ${toFill.length} field(s). Generating answers...`, 'ai');

        const withAnswers = await generateAnswers(toFill, companyProfile, resolvedName, sectionLabel, sectionNum, filledContext);

        io.emit('fill_start', {
          section: sectionNum, heading: sectionLabel, total: withAnswers.length,
          fields: withAnswers.map(f => ({ label: f.label, type: f.type, purpose: f.purpose, answer: f.answer }))
        });
        io.emit('status', 'filling');

        const results = [];
        const contextLines = [];

        for (let i = 0; i < withAnswers.length; i++) {
          if (!agentRunning) break;
          const field = withAnswers[i];
          io.emit('filling_field', { index: i, label: field.label, purpose: field.purpose, answer: field.answer, section: sectionNum });
          log(`[${field.purpose}] "${field.label}" → "${String(field.answer).slice(0,60)}"`, 'info');

          const result = await fillField(activePage, field);
          results.push({ ...field, result });
          knownSelectors.add(field.selector);
          totalFields++;

          if (result.status === 'filled') {
            totalFilled++;
            contextLines.push(`${field.label}: ${result.value}`);
            log(`✓ "${field.label}" — ${result.value?.toString().slice(0,50)}`, 'success');
          } else if (result.status === 'skipped') {
            totalSkipped++;
            log(`— "${field.label}" skipped: ${result.reason}`, 'warn');
          } else {
            totalFailed++;
            log(`✗ "${field.label}": ${result.reason||result.status}`, 'warn');
          }
          io.emit('field_done', { index: i, status: result.status, value: result.value, reason: result.reason, section: sectionNum });

          // Check for newly appeared conditional fields after each fill
          await page.waitForTimeout(400);
          const nowFields = await scanFields(activePage);
          const appeared = nowFields.filter(f => !knownSelectors.has(f.selector));
          if (appeared.length > 0) {
            log(`${appeared.length} conditional field(s) appeared`, 'info');
            io.emit('conditional_fields', { count: appeared.length, fields: appeared.map(f=>f.label) });
            const condWithAnswers = await generateAnswers(appeared, companyProfile, resolvedName, sectionLabel, sectionNum, filledContext + '\n' + contextLines.join('\n'));
            for (const cf of condWithAnswers) {
              if (!agentRunning) break;
              knownSelectors.add(cf.selector);
              io.emit('filling_field', { index: i+0.5, label: cf.label, purpose: cf.purpose, answer: cf.answer, section: sectionNum });
              log(`[conditional][${cf.purpose}] "${cf.label}" → "${String(cf.answer).slice(0,60)}"`, 'info');
              const cr = await fillField(activePage, cf);
              results.push({ ...cf, result: cr });
              totalFields++;
              if (cr.status === 'filled') { totalFilled++; contextLines.push(`${cf.label}: ${cr.value}`); }
              else if (cr.status === 'skipped') totalSkipped++;
              else totalFailed++;
              io.emit('field_done', { index: i+0.5, status: cr.status, value: cr.value, reason: cr.reason, section: sectionNum });
            }
          }
        }

        filledContext += (filledContext ? '\n' : '') + contextLines.join('\n');
        const filled = results.filter(r => r.result?.status==='filled').length;
        log(`Section ${sectionNum} done: ${filled}/${toFill.length} filled`, 'success');
        io.emit('section_filled', { num: sectionNum, filled, total: toFill.length, results: results.map(r=>({ label: r.label, purpose: r.purpose, answer: r.answer, status: r.result?.status, reason: r.result?.reason })) });
      }

      if (!agentRunning) break;

      // Pause for human review
      const action = await waitForHuman(sectionLabel);
      if (action === 'stop' || !agentRunning) break;

      // Proceed to next section
      const prevHeading = state.heading;
      const prevUrl = state.url;
      log('Proceeding to next section...', 'system');
      const btnText = await clickProceedButton(activePage);

      if (btnText) {
        log(`Clicked: "${btnText}"`, 'success');
        io.emit('status', 'navigating');
        await activePage.waitForTimeout(2000);
        try { await activePage.waitForLoadState('domcontentloaded', { timeout: 8000 }); } catch {}

        const newState = await waitForChange(activePage, prevHeading, prevUrl, 8000);
        if (!newState) {
          log('Page did not change — possible validation error. Fix in Chrome then Continue.', 'warn');
          io.emit('page_stuck');
          await waitForHuman('Fix validation errors then continue');
          if (!agentRunning) break;
        } else if (newState.isConfirmation) {
          log('Form submitted! Confirmation detected.', 'success');
          io.emit('form_submitted');
          updateSheet(resolvedName, 'submitted', totalFilled, totalFields, `${totalFilled}/${totalFields} fields — submitted`);
          break;
        }
      } else {
        log('No proceed button found — click it manually in Chrome, then Continue here.', 'warn');
        io.emit('no_next_btn');
        await waitForHuman('Click the button manually in Chrome');
        if (!agentRunning) break;
        await activePage.waitForTimeout(2000);
      }
    }

    // Update sheet when finished
    if (agentRunning) {
      updateSheet(resolvedName, 'filled', totalFilled, totalFields, `${totalFilled}/${totalFields} fields filled — review and submit`);
    }
    log(`Done. Filled: ${totalFilled}, Skipped: ${totalSkipped}, Failed: ${totalFailed}`, 'success');
    io.emit('agent_done_final', { totalFilled, totalSkipped, totalFailed, totalFields });

  } catch(err) {
    log(`Error: ${err.message}`, 'error');
    updateSheet(resolvedName, 'error', totalFilled, totalFields, err.message.slice(0,100));
    io.emit('agent_error', { error: err.message });
  } finally {
    agentRunning = false;
    humanConfirmResolve = null;
  }
});

app.post('/api/close', async (req, res) => {
  if (activeBrowser) { await activeBrowser.close(); activeBrowser = null; activePage = null; }
  res.json({ closed: true });
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`\nGrantAgent → http://localhost:${PORT}`);
  console.log(`Model: ${OLLAMA_MODEL} | Ollama: ${OLLAMA_URL}\n`);
});
