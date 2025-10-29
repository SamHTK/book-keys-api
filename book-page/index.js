module.exports = async function (context, req) {
  context.log("=== Handler started ===");
  
  try {
    // Move the require INSIDE the handler
    context.log("Attempting to load pages module...");
    const { getPageConfig } = require("../lib/pages");
    context.log("Pages module loaded successfully");
    
    const slug = context.bindingData.slug;
    context.log("Processing slug:", slug);
    
    const cfg = getPageConfig(slug);
    context.log("Config loaded:", JSON.stringify({ slug: cfg.slug }));
    
    context.res = { 
      status: 200, 
      headers: { 'Content-Type': 'text/html' }, 
      body: pageHtml(cfg) 
    };
    
  } catch (err) {
    context.log.error("=== ERROR ===");
    context.log.error("Error type:", err.constructor.name);
    context.log.error("Error message:", err.message);
    context.log.error("Error stack:", err.stack);
    
    context.res = { 
      status: 404, 
      headers: { 'Content-Type': 'text/html' }, 
      body: `<!doctype html><html><body><p>Error: ${err.message}</p></body></html>` 
    };
  }
};

function htmlEscape(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function pageHtml({ slug, schedulerUpn, timeZone, businessHours }) {
  const title = `Book time with ${schedulerUpn}`;
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${htmlEscape(title)}</title>
  <style>
    body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f8fafc;color:#0f172a}
    header{padding:16px 20px;background:#0f172a;color:#fff}
    main{max-width:960px;margin:0 auto;padding:20px}
    .row{display:flex;gap:20px;flex-wrap:wrap}
    .col{flex:1 1 320px}
    .card{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:16px}
    .label{font-weight:600;margin:6px 0 4px}
    input,select,textarea{width:100%;padding:10px;border:1px solid #cbd5e1;border-radius:6px}
    textarea{min-height:90px}
    .slots{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:8px}
    .slot{padding:10px;border:1px solid #cbd5e1;border-radius:6px;text-align:center;cursor:pointer;background:#fff}
    .slot:hover{background:#f1f5f9}
    .slot.selected{background:#2563eb;color:#fff;border-color:#2563eb}
    .actions{display:flex;gap:12px;align-items:center;margin-top:12px}
    .btn{background:#2563eb;color:#fff;padding:10px 16px;border-radius:6px;border:none;cursor:pointer}
    .btn:disabled{opacity:.6;cursor:not-allowed}
    .muted{color:#64748b}
    .error{color:#b91c1c}
    .success{color:#065f46}
  </style>
</head>
<body>
  <header>
    <h2 style="margin:0">${htmlEscape(title)}</h2>
    <div class="muted">Time zone: ${htmlEscape(timeZone)} | Business hours: ${htmlEscape(businessHours.start)}–${htmlEscape(businessHours.end)}</div>
  </header>
  <main>
    <div class="row">
      <div class="col card">
        <div class="label">Date</div>
        <input id="date" type="date" />
        <div class="label">Duration</div>
        <select id="duration">
          <option value="15">15 minutes</option>
          <option value="30" selected>30 minutes</option>
          <option value="45">45 minutes</option>
          <option value="60">60 minutes</option>
        </select>
        <div class="label">Available slots</div>
        <div id="slots" class="slots"></div>
        <div id="slotsStatus" class="muted" style="margin-top:8px"></div>
      </div>
      <div class="col card">
        <div class="label">Your email</div>
        <input id="email" type="email" placeholder="you@example.com" />
        <div class="label">Title</div>
        <input id="title" type="text" placeholder="Meeting subject" />
        <div class="label">Additional attendees (comma-separated emails, optional)</div>
        <input id="attendees" type="text" placeholder="alice@example.com,bob@example.com" />
        <div class="label"><label><input id="wantsTeams" type="checkbox" /> Request Microsoft Teams meeting</label></div>
        <div class="label">Notes (optional)</div>
        <textarea id="notes" placeholder="Add any context"></textarea>
        <div class="actions">
          <button id="submit" class="btn" disabled>Request booking</button>
          <span id="submitStatus" class="muted"></span>
        </div>
      </div>
    </div>
  </main>
  <script>
  const slug = ${JSON.stringify(slug)};
  const tz = ${JSON.stringify(timeZone)};

// Convert Windows timezone names to IANA timezone names
function windowsToIANA(windowsTz) {
  const mapping = {
    'Eastern Standard Time': 'America/New_York',
    'Central Standard Time': 'America/Chicago',
    'Mountain Standard Time': 'America/Denver',
    'Pacific Standard Time': 'America/Los_Angeles',
    'GMT Standard Time': 'Europe/London',
    'Central European Standard Time': 'Europe/Paris',
    'Tokyo Standard Time': 'Asia/Tokyo',
    'AUS Eastern Standard Time': 'Australia/Sydney',
  };
  return mapping[windowsTz] || windowsTz;
}

function fmtLocal(dtStr) {
  try {
    // Parse the ISO datetime string (it's in UTC)
    const dt = new Date(dtStr);
    const ianaTz = windowsToIANA(tz);
    
    // Format in the configured timezone
    return dt.toLocaleTimeString('en-US', { 
      timeZone: ianaTz,
      hour: 'numeric', 
      minute: '2-digit', 
      hour12: true 
    });
  } catch {
    return dtStr;
  }
}

function ymd(d) {
  // Get the date in the configured timezone
  const ianaTz = windowsToIANA(tz);
  const parts = d.toLocaleString('en-US', { timeZone: ianaTz }).split(',')[0].split('/');
  const month = parts[0].padStart(2,'0');
  const day = parts[1].padStart(2,'0');
  const year = parts[2];
  return \`\${year}-\${month}-\${day}\`;
}

  const dateEl = document.getElementById('date');
  const durationEl = document.getElementById('duration');
  const slotsEl = document.getElementById('slots');
  const slotsStatusEl = document.getElementById('slotsStatus');
  const submitBtn = document.getElementById('submit');
  const submitStatus = document.getElementById('submitStatus');

  const emailEl = document.getElementById('email');
  const titleEl = document.getElementById('title');
  const attendeesEl = document.getElementById('attendees');
  const wantsTeamsEl = document.getElementById('wantsTeams');
  const notesEl = document.getElementById('notes');

  let selected = null;

  function validateForm() {
    const ok = !!selected && emailEl.value.trim() && titleEl.value.trim();
    submitBtn.disabled = !ok;
  }

  function renderSlots(slots) {
    slotsEl.innerHTML = '';
    selected = null;
    validateForm();
    if (!slots || !slots.length) {
      slotsStatusEl.textContent = 'No slots available on this day.';
      return;
    }
    slotsStatusEl.textContent = '';
    for (const s of slots) {
      const btn = document.createElement('button');
      btn.className = 'slot';
      btn.type = 'button';
      btn.textContent = fmtLocal(s.start);
      btn.addEventListener('click', () => {
        document.querySelectorAll('.slot.selected').forEach(el => el.classList.remove('selected'));
        btn.classList.add('selected');
        selected = s;
        validateForm();
      });
      slotsEl.appendChild(btn);
    }
  }

  async function loadSlots() {
    slotsEl.innerHTML = '';
    slotsStatusEl.textContent = 'Loading...';
    selected = null;
    validateForm();
    const d = dateEl.value;
    const duration = durationEl.value;
    try {
      const resp = await fetch('/api/book/' + encodeURIComponent(slug) + '/slots?date=' + encodeURIComponent(d) + '&duration=' + encodeURIComponent(duration));
      const json = await resp.json();
      if (!resp.ok) throw new Error(json && json.error || 'Failed to load slots');
      renderSlots(json.slots || []);
    } catch (e) {
      slotsEl.innerHTML = '';
      slotsStatusEl.textContent = 'Error loading slots: ' + e.message;
    }
  }

  dateEl.addEventListener('change', loadSlots);
  durationEl.addEventListener('change', loadSlots);
  emailEl.addEventListener('input', validateForm);
  titleEl.addEventListener('input', validateForm);

  submitBtn.addEventListener('click', async () => {
    if (!selected) return;
    submitBtn.disabled = true;
    submitStatus.textContent = 'Submitting...';
    try {
      const duration = Number(durationEl.value);
      const body = {
        start: selected.start,
        duration,
        title: titleEl.value.trim(),
        email: emailEl.value.trim(),
        additionalAttendees: attendeesEl.value.split(',').map(s => s.trim()).filter(Boolean),
        wantsTeams: !!wantsTeamsEl.checked,
        notes: notesEl.value
      };
      const resp = await fetch('/api/book/' + encodeURIComponent(slug) + '/request', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      const json = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(json && json.error || 'Failed to submit request');
      submitStatus.className = 'success';
      submitStatus.textContent = 'Request submitted. Check your email for confirmation after approval.';
    } catch (e) {
      submitStatus.className = 'error';
      submitStatus.textContent = e.message;
    } finally {
      submitBtn.disabled = false;
    }
  });

  const today = new Date();
  dateEl.value = ymd(today);
  loadSlots();
  </script>
</body>
</html>`;
}
