document.querySelectorAll('.slot input[type=file]').forEach(function (input) {
  input.addEventListener('change', function () {
    const slot = input.closest('.slot');
    const nameBox = slot.querySelector('.filename');
    if (input.files && input.files[0]) {
      slot.classList.add('filled');
      nameBox.textContent = input.files[0].name;
    } else {
      slot.classList.remove('filled');
      nameBox.textContent = '';
    }
  });
});

const form = document.getElementById('recon-form');
const runBtn = document.getElementById('run-btn');
const progress = document.getElementById('progress');
const resultsBox = document.getElementById('results');

function badgeClass(status) {
  if (status === 'Matched') return 'matched';
  if (status === 'Pending - Multiple Possible Matches') return 'conflict';
  if (status.indexOf('Not Applicable') === 0) return 'na';
  return 'pending';
}

function renderResults(runId, summary) {
  let tilesHtml = `
    <div class="tiles">
      <div class="tile"><span class="num">${summary.total}</span><span class="lbl">Total Entries</span></div>
      <div class="tile matched"><span class="num">${summary.matched}</span><span class="lbl">Matched</span></div>
      <div class="tile pending"><span class="num">${summary.pending}</span><span class="lbl">Pending Review</span></div>
      <div class="tile"><span class="num">${summary.hdfc_unmatched}</span><span class="lbl">HDFC Lines Unmatched</span></div>
    </div>`;

  let rowsHtml = '';
  const order = ['3493', '3496', '345051'];
  order.forEach(function (gl) {
    const s = summary.per_gl[gl];
    if (!s) return;
    const pct = s.total ? Math.round((s.matched / s.total) * 100) : 0;
    rowsHtml += `
      <tr>
        <td><strong>${gl}</strong> — ${s.description}</td>
        <td class="num">${s.total}</td>
        <td class="num" style="color: var(--verified)">${s.matched}</td>
        <td class="num" style="color: var(--query)">${s.pending}</td>
        <td class="num">${s.na}</td>
        <td class="num">${pct}%</td>
      </tr>`;
  });

  const overallOk = summary.pending === 0;
  const stamp = overallOk
    ? '<span class="stamp ok">✓ Fully Reconciled</span>'
    : `<span class="stamp warn">${summary.pending} Pending Review</span>`;

  resultsBox.innerHTML = `
    <div class="card">
      <h2>Result ${stamp}</h2>
      ${tilesHtml}
      <table class="ledger">
        <thead>
          <tr><th>GL Head</th><th class="num">Total</th><th class="num">Matched</th><th class="num">Pending</th><th class="num">N/A</th><th class="num">Match %</th></tr>
        </thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <div class="download-row">
        <a class="btn" href="/runs/${runId}">View Matched &amp; Pending Entries</a>
        <a class="btn secondary" href="/reports/${runId}">Download Full Excel Report</a>
        <a class="btn secondary" href="/history">View in History</a>
      </div>
    </div>`;
}

form.addEventListener('submit', async function (e) {
  e.preventDefault();
  runBtn.disabled = true;
  progress.style.display = 'block';
  resultsBox.innerHTML = '';

  const formData = new FormData(form);
  try {
    const resp = await fetch('/api/reconcile', { method: 'POST', body: formData });
    const data = await resp.json();
    if (!data.ok) {
      resultsBox.innerHTML = `<div class="card"><div class="error-banner">Reconciliation failed: ${data.error}</div></div>`;
    } else {
      renderResults(data.run_id, data.summary);
    }
  } catch (err) {
    resultsBox.innerHTML = `<div class="card"><div class="error-banner">Request failed: ${err}</div></div>`;
  } finally {
    runBtn.disabled = false;
    progress.style.display = 'none';
  }
});
