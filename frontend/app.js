/**
 * AI Velocity Portal - Dashboard with Leaderboard + Compact Lists
 */
const API = window.location.origin + '/api';
let allSubmissions = [], allCompleted = [], allInProgress = [], allPendingIdeas = [];

document.addEventListener('DOMContentLoaded', () => {
    initNavbar(); initNeuralNetwork(); initMobileToggle(); loadDashboard();
});

// ─── Load All Data ─────────────────────────────────────────────────────────────
async function loadDashboard() {
    try {
        const [stats, subs, leaders, procs, stages] = await Promise.all([
            fetch(`${API}/stats`).then(r => r.json()),
            fetch(`${API}/submissions`).then(r => r.json()),
            fetch(`${API}/leaderboard`).then(r => r.json()),
            fetch(`${API}/processes`).then(r => r.json()),
            fetch(`${API}/stages`).then(r => r.json()),
        ]);
        allSubmissions = subs;
        allCompleted = subs.filter(s => s.category === 'ai_win');
        allInProgress = subs.filter(s => s.category === 'new_idea' && (s.status === 'Approved' || s.status === 'In Review'));
        allPendingIdeas = subs.filter(s => s.category === 'new_idea' && s.status !== 'Approved' && s.status !== 'In Review');

        renderKPIs(stats);
        renderStagesChart(stages);
        renderStatusChart(stats.statuses);
        renderLeaderboard(leaders);
        renderCompletedSplit(allCompleted);
        renderInProgressList(allInProgress);
        renderIdeasList(allPendingIdeas);
        populateFilters(leaders, procs);
        initFilterListeners();
        initScrollAnimations();
        showToast(`✅ Loaded ${subs.length} submissions`);
    } catch (err) {
        console.error(err);
        showToast('⚠️ Failed to load data');
    }
}

async function refreshData() {
    showToast('🔄 Refreshing data...');
    await loadDashboard();
}

async function syncFromSharePoint() {
    showToast('📡 Starting SharePoint sync...');
    try {
        const resp = await fetch(`${API}/sync`, { method: 'POST' });
        const data = await resp.json();
        if (data.status === 'started') {
            showToast('🚀 Sync started! Edge browser will open to download CSV. Dashboard will auto-refresh in 60s.');
            setTimeout(() => { showToast('🔄 Auto-refreshing data...'); loadDashboard(); }, 60000);
        } else {
            showToast('⚠️ Sync failed: ' + (data.message || 'Unknown error'));
        }
    } catch (err) {
        showToast('⚠️ Could not start sync: ' + err.message);
    }
}

// ─── KPIs ──────────────────────────────────────────────────────────────────────
function renderKPIs(stats, inProgressCount) {
    animateValue('kpiTotal', stats.total_submissions);
    animateValue('kpiProd', stats.live_in_production);
    animateValue('kpiUat', stats.uat_in_progress);
    animateValue('kpiIdeas', stats.new_ideas);
    animateValue('kpiRepl', stats.replicate_requests);
    animateValue('kpiPeople', stats.unique_submitters);
}
function animateValue(id, target) {
    const el = document.getElementById(id); if (!el) return;
    const dur = 1200, start = performance.now();
    (function update(now) {
        const p = Math.min((now - start) / dur, 1);
        el.textContent = Math.floor(target * (1 - Math.pow(1 - p, 3)));
        if (p < 1) requestAnimationFrame(update);
    })(start);
}

// ─── Charts ────────────────────────────────────────────────────────────────────
function renderStagesChart(stages) {
    const c = document.getElementById('processChart'); if (!c) return;
    const entries = Object.entries(stages);
    if (!entries.length) { c.innerHTML = '<div class="empty-state">No implementation stage data</div>'; return; }
    const total = entries.reduce((a, [,v]) => a + v, 0);
    const max = Math.max(...entries.map(([,v]) => v));
    const icons = { 'In Progress (Development Stage)': '🔧', 'In Progress (UAT Stage)': '🧪', 'Completed (Awaiting Approvals)': '⏳', 'Completed (Production)': '✅' };
    const colors = { 'In Progress (Development Stage)': '#6366f1', 'In Progress (UAT Stage)': '#f59e0b', 'Completed (Awaiting Approvals)': '#8b5cf6', 'Completed (Production)': '#10b981' };
    const shortNames = { 'In Progress (Development Stage)': 'Development', 'In Progress (UAT Stage)': 'UAT / Testing', 'Completed (Awaiting Approvals)': 'Awaiting Approvals', 'Completed (Production)': 'Live in Production' };
    c.innerHTML = entries.map(([name, v]) => {
        const pct = ((v/total)*100).toFixed(0);
        const icon = icons[name] || '📋';
        const color = colors[name] || '#6b6b80';
        const label = shortNames[name] || name;
        return `<div class="stage-row">
            <div class="stage-icon">${icon}</div>
            <div class="stage-info">
                <div class="stage-name">${label}</div>
                <div class="stage-bar-track"><div class="stage-bar-fill" style="width:${(v/max)*100}%;background:${color}"></div></div>
            </div>
            <div class="stage-count" style="color:${color}">${v} <span class="stage-pct">(${pct}%)</span></div>
        </div>`;
    }).join('') + `<div class="stage-total">Total active implementations: <strong>${total}</strong></div>`;
}
function renderStatusChart(statuses) {
    const c = document.getElementById('statusChart'); if (!c) return;
    const colors = { Completed:'#10b981', Approved:'#6366f1', 'In Review':'#f59e0b', Pending:'#8b5cf6', New:'#06b6d4' };
    const total = Object.values(statuses).reduce((a,b) => a+b, 0);
    c.innerHTML = Object.entries(statuses).map(([s, v]) => {
        const pct = ((v/total)*100).toFixed(1), color = colors[s] || '#6b6b80';
        return `<div class="status-row"><div class="status-dot" style="background:${color}"></div>
        <div class="status-name">${s}</div><div class="status-bar-track"><div class="status-bar-fill" style="width:${pct}%;background:${color}"></div></div>
        <div class="status-count">${v} (${pct}%)</div></div>`;
    }).join('');
}

// ─── Leaderboard ───────────────────────────────────────────────────────────────
function renderLeaderboard(leaders) {
    const c = document.getElementById('leaderboardWrap'); if (!c) return;
    const medals = ['🥇','🥈','🥉','4️⃣','5️⃣','6️⃣','7️⃣'];
    const totalAll = leaders.reduce((a, l) => a + l.total, 0);
    c.innerHTML = `
    <table class="leader-table">
        <thead><tr>
            <th>#</th><th>LEADER</th><th>POC</th><th>TOTAL</th>
            <th>🏆 WINS</th><th>⚡ IN PROG</th><th>💡 IDEAS</th><th>🔁 REPL</th>
            <th>👥 PEOPLE</th><th>SHARE</th><th>PROGRESS</th>
        </tr></thead>
        <tbody>
        ${leaders.map((l, i) => {
            const share = ((l.total / totalAll) * 100).toFixed(0);
            return `<tr class="leader-row" onclick="filterByLeader('${l.leader}')">
                <td class="rank">${medals[i] || i+1}</td>
                <td class="leader-name">${l.leader}</td>
                <td class="leader-poc">${l.poc}</td>
                <td class="num bold">${l.total}</td>
                <td class="num wins">${l.wins}</td>
                <td class="num progress">${l.in_progress}</td>
                <td class="num ideas">${l.ideas}</td>
                <td class="num repl">${l.replicates}</td>
                <td class="num">${l.contributors}</td>
                <td class="num">${share}%</td>
                <td><div class="leader-bar-track"><div class="leader-bar-fill" style="width:${share}%"></div></div></td>
            </tr>`;
        }).join('')}
        </tbody>
        <tfoot><tr>
            <td></td><td class="bold">TOTAL</td><td></td>
            <td class="num bold">${totalAll}</td>
            <td class="num bold">${leaders.reduce((a,l) => a+l.wins, 0)}</td>
            <td class="num bold">${leaders.reduce((a,l) => a+l.in_progress, 0)}</td>
            <td class="num bold">${leaders.reduce((a,l) => a+l.ideas, 0)}</td>
            <td class="num bold">${leaders.reduce((a,l) => a+l.replicates, 0)}</td>
            <td class="num bold" title="Unique across all orgs">${leaders.reduce((a,l) => a+l.contributors, 0)}*</td>
            <td class="num bold">100%</td><td></td>
        </tr></tfoot>
    </table>`;
}

function filterByLeader(leader) {
    document.getElementById('leaderFilter').value = leader;
    applyIdeasFilter();
    document.getElementById('ideas')?.scrollIntoView({ behavior: 'smooth' });
}

// ─── Compact List: Completed Projects (Split into Production + UAT/Pending) ───
function renderCompletedSplit(items) {
    // "True Complete" = ONLY those explicitly marked "Completed (Production)" i.e. "Not required - its a completed win ready for production"
    const prodItems = items.filter(w => w.implementation_stage === 'Completed (Production)');
    // "UAT/Approval Pending" = everything else (Awaiting Approvals, UAT, Development, or empty/unknown)
    const uatItems = items.filter(w => w.implementation_stage !== 'Completed (Production)');

    // Update counts
    const prodCountEl = document.getElementById('prodCount');
    const uatCountEl = document.getElementById('uatCount');
    if (prodCountEl) prodCountEl.textContent = `(${prodItems.length})`;
    if (uatCountEl) uatCountEl.textContent = `(${uatItems.length})`;

    // Render Production table
    renderCompletedTable('completedProdList', prodItems);
    // Render UAT/Pending table
    renderCompletedTable('completedUatList', uatItems);
}

function renderCompletedTable(containerId, items) {
    const c = document.getElementById(containerId); if (!c) return;
    if (!items.length) { c.innerHTML = '<div class="empty-state">No projects in this category</div>'; return; }
    c.innerHTML = `<table class="data-table">
        <thead><tr><th>#</th><th>Project</th><th>Owner</th><th>Process</th><th>Leader</th><th>Impact</th><th>Stage</th><th>Replicable</th></tr></thead>
        <tbody>${items.map((w, i) => {
            const stageLabel = w.implementation_stage ? w.implementation_stage.replace('Completed (','').replace('In Progress (','').replace(')','') : 'Production';
            return `<tr class="data-row clickable" onclick="openDetail(${w.id})">
                <td>${i+1}</td>
                <td class="col-name">${w.project_name || 'Untitled'}</td>
                <td>${w.project_owner || w.name}</td>
                <td class="col-process">${trunc(w.process, 25)}</td>
                <td><span class="leader-chip">${w.leader}</span></td>
                <td class="col-impact">${trunc(w.impact, 50)}</td>
                <td><span class="stage-chip">${stageLabel}</span></td>
                <td>${w.replicable === 'Yes' ? '✅' : '—'}</td>
            </tr>`;
        }).join('')}
        </tbody></table>`;
}

// ─── Compact List: In Progress ─────────────────────────────────────────────────
function renderInProgressList(items) {
    const c = document.getElementById('inprogressList'); if (!c) return;
    if (!items.length) { c.innerHTML = '<div class="empty-state">No in-progress items</div>'; return; }
    c.innerHTML = `<table class="data-table">
        <thead><tr><th>#</th><th>Submitter</th><th>Process</th><th>Leader</th><th>Status</th><th>Problem</th><th>Proposed Solution</th><th>Timeline</th><th>SDE</th></tr></thead>
        <tbody>${items.map((it, i) => `
            <tr class="data-row clickable" onclick="openDetail(${it.id})">
                <td>${i+1}</td>
                <td>${it.name || 'Anonymous'}</td>
                <td class="col-process">${trunc(it.process, 22)}</td>
                <td><span class="leader-chip">${it.leader}</span></td>
                <td><span class="status-pill status-${it.status.toLowerCase().replace(' ','-')}">${it.status}</span></td>
                <td class="col-problem">${trunc(it.problem_statement, 60)}</td>
                <td class="col-solution">${trunc(it.proposed_solution || it.ai_solution, 50)}</td>
                <td>${it.target_timeline || '—'}</td>
                <td class="col-sde">${it.sde_contact?.name ? trunc(it.sde_contact.name,15) : 'TBD'}</td>
            </tr>`).join('')}
        </tbody></table>`;
}

// ─── Compact List: Submitted Ideas ─────────────────────────────────────────────
function renderIdeasList(items) {
    const c = document.getElementById('ideasList'); if (!c) return;
    if (!items.length) { c.innerHTML = '<div class="empty-state">No ideas match your filters</div>'; return; }
    c.innerHTML = `<table class="data-table">
        <thead><tr><th>#</th><th>Submitter</th><th>Process</th><th>Leader</th><th>Problem Statement</th><th>Proposed AI Solution</th><th>Effort</th><th>Tools</th><th>SDE</th></tr></thead>
        <tbody>${items.map((it, i) => `
            <tr class="data-row clickable" onclick="openDetail(${it.id})">
                <td>${i+1}</td>
                <td>${it.name || 'Anonymous'}</td>
                <td class="col-process">${trunc(it.process, 22)}</td>
                <td><span class="leader-chip">${it.leader}</span></td>
                <td class="col-problem">${trunc(it.problem_statement, 70)}</td>
                <td class="col-solution">${trunc(it.proposed_solution || it.ai_solution, 50)}</td>
                <td>${it.current_effort || '—'}</td>
                <td class="col-tools">${(it.suggested_tools||[]).slice(0,2).map(t=>`<span class="tool-tag-sm">${t}</span>`).join('')}</td>
                <td class="col-sde">${it.sde_contact?.name ? trunc(it.sde_contact.name,12) : 'TBD'}</td>
            </tr>`).join('')}
        </tbody></table>`;
}

// ─── Filters ───────────────────────────────────────────────────────────────────
function populateFilters(leaders, processes) {
    const lf = document.getElementById('leaderFilter');
    const pf = document.getElementById('processFilter');
    leaders.forEach(l => { const o = document.createElement('option'); o.value = l.leader; o.textContent = l.leader; lf?.appendChild(o); });
    processes.forEach(p => { const o = document.createElement('option'); o.value = p; o.textContent = p.length > 35 ? p.substring(0,33)+'...' : p; pf?.appendChild(o); });
}
function initFilterListeners() {
    document.getElementById('ideaSearch')?.addEventListener('input', debounce(applyIdeasFilter, 300));
    document.getElementById('leaderFilter')?.addEventListener('change', applyIdeasFilter);
    document.getElementById('processFilter')?.addEventListener('change', applyIdeasFilter);
}
function applyIdeasFilter() {
    let filtered = [...allPendingIdeas];
    const q = (document.getElementById('ideaSearch')?.value || '').toLowerCase();
    const leader = document.getElementById('leaderFilter')?.value || '';
    const proc = document.getElementById('processFilter')?.value || '';
    if (q) filtered = filtered.filter(i =>
        (i.name||'').toLowerCase().includes(q) || (i.problem_statement||'').toLowerCase().includes(q) ||
        (i.proposed_solution||'').toLowerCase().includes(q) || (i.process||'').toLowerCase().includes(q));
    if (leader) filtered = filtered.filter(i => i.leader === leader);
    if (proc) filtered = filtered.filter(i => i.process === proc);
    renderIdeasList(filtered);
}

// ─── Detail Modal ──────────────────────────────────────────────────────────────
function openDetail(id) {
    const item = allSubmissions.find(s => s.id === id); if (!item) return;
    const modal = document.getElementById('modalOverlay');
    const content = document.getElementById('modalContent');
    const cat = { ai_win:'🏆 Completed AI Win', new_idea:'💡 AI Idea', replicate:'🔁 Replicate' }[item.category] || '';
    const sc = item.status.toLowerCase().replace(' ','-');
    content.innerHTML = `
        <div class="modal-header"><span class="modal-category">${cat}</span><span class="status-pill status-${sc}">${item.status}</span></div>
        <h2 class="modal-title">${item.project_name || trunc(item.problem_statement, 80) || 'Submission #'+item.id}</h2>
        <div class="modal-info-grid">
            <div class="modal-info-item"><span class="info-label">👤 Submitted By</span><span class="info-value">${item.name || item.created_by}</span></div>
            <div class="modal-info-item"><span class="info-label">🏢 Process</span><span class="info-value">${item.process}</span></div>
            <div class="modal-info-item"><span class="info-label">📂 Sub Process</span><span class="info-value">${item.sub_process || 'N/A'}</span></div>
            <div class="modal-info-item"><span class="info-label">🏅 Leader</span><span class="info-value">${item.leader}</span></div>
            <div class="modal-info-item"><span class="info-label">👔 Manager</span><span class="info-value">${item.manager || 'N/A'}</span></div>
            <div class="modal-info-item"><span class="info-label">📅 Created</span><span class="info-value">${item.created || 'N/A'}</span></div>
        </div>
        ${item.category === 'ai_win' ? `
            <div class="modal-section"><h3>🎯 Challenge</h3><p>${item.challenge || 'N/A'}</p></div>
            <div class="modal-section"><h3>🤖 AI Solution</h3><p>${item.ai_solution || 'N/A'}</p></div>
            <div class="modal-section"><h3>📈 Impact</h3><p>${item.impact || 'N/A'}</p></div>
            ${item.project_team ? `<div class="modal-section"><h3>👥 Team</h3><p>${item.project_team}</p></div>` : ''}
            ${item.replicable ? `<div class="modal-section"><h3>🔄 Replicable?</h3><p>${item.replicable}</p></div>` : ''}
        ` : ''}
        ${item.category === 'new_idea' ? `
            <div class="modal-section"><h3>❗ Problem</h3><p>${item.problem_statement || 'N/A'}</p></div>
            <div class="modal-section"><h3>💡 Proposed Solution</h3><p>${item.proposed_solution || item.ai_solution || 'N/A'}</p></div>
            ${item.current_effort ? `<div class="modal-section"><h3>⏱️ Manual Effort</h3><p>${item.current_effort}</p></div>` : ''}
            ${item.estimated_volume ? `<div class="modal-section"><h3>📊 Volume</h3><p>${item.estimated_volume}</p></div>` : ''}
            ${item.execution_plan ? `<div class="modal-section"><h3>🔧 Execution Plan</h3><p>${item.execution_plan}</p></div>` : ''}
        ` : ''}
        ${item.category === 'replicate' ? `
            <div class="modal-section"><h3>🔁 Win to Replicate</h3><p>${item.which_win_to_replicate || 'N/A'}</p></div>
            <div class="modal-section"><h3>📋 Current Process</h3><p>${item.current_process_desc || 'N/A'}</p></div>
        ` : ''}
        <div class="modal-section"><h3>🤖 Suggested AI Tools</h3><div class="tools-list">${(item.suggested_tools||[]).map(t=>`<span class="tool-tag">${t}</span>`).join('')}</div></div>
        <div class="modal-section"><h3>👨‍💻 SDE Contact</h3>
            <div class="sde-card"><div class="sde-avatar">👨‍💻</div><div>
                <div class="sde-name-full">${item.sde_contact?.name || 'TBD'}</div>
                <div class="sde-alias">${item.sde_contact?.alias ? item.sde_contact.alias+'@amazon.com' : ''}</div>
            </div></div>
        </div>`;
    modal.style.display = 'flex'; document.body.style.overflow = 'hidden';
}
function closeModal() { document.getElementById('modalOverlay').style.display = 'none'; document.body.style.overflow = ''; }
document.getElementById('modalOverlay')?.addEventListener('click', e => { if (e.target.id === 'modalOverlay') closeModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

// ─── Utilities ─────────────────────────────────────────────────────────────────
function trunc(s, l) { return !s ? '' : s.length > l ? s.substring(0, l) + '...' : s; }
function debounce(fn, d) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), d); }; }
function showToast(msg) { const t = document.getElementById('toast'), m = document.getElementById('toastMessage'); if (!t||!m) return; m.textContent = msg; t.classList.add('show'); setTimeout(() => t.classList.remove('show'), 3000); }

// ─── Navbar ────────────────────────────────────────────────────────────────────
function initNavbar() {
    window.addEventListener('scroll', () => {
        const nb = document.getElementById('navbar');
        nb?.classList.toggle('scrolled', window.scrollY > 50);
        document.querySelectorAll('section[id]').forEach(s => {
            const r = s.getBoundingClientRect();
            if (r.top <= 200 && r.bottom >= 200) {
                document.querySelectorAll('.nav-link').forEach(l => { l.classList.toggle('active', l.getAttribute('href') === '#' + s.id); });
            }
        });
    });
}
function initMobileToggle() {
    document.getElementById('mobileToggle')?.addEventListener('click', () => {
        const nl = document.querySelector('.nav-links'); nl.style.display = nl.style.display === 'flex' ? 'none' : 'flex';
    });
}

// ─── Neural Net BG ─────────────────────────────────────────────────────────────
function initNeuralNetwork() {
    const el = document.getElementById('neuralNetwork'); if (!el) return;
    const cvs = document.createElement('canvas'); cvs.style.cssText = 'width:100%;height:100%'; el.appendChild(cvs);
    const ctx = cvs.getContext('2d'); let particles = [];
    function resize() { cvs.width = window.innerWidth; cvs.height = window.innerHeight; } resize(); window.addEventListener('resize', resize);
    class P { constructor() { this.x=Math.random()*cvs.width; this.y=Math.random()*cvs.height; this.vx=(Math.random()-0.5)*0.4; this.vy=(Math.random()-0.5)*0.4; this.r=Math.random()*2+1; this.o=Math.random()*0.4+0.2; }
        update() { this.x+=this.vx; this.y+=this.vy; if(this.x<0||this.x>cvs.width) this.vx*=-1; if(this.y<0||this.y>cvs.height) this.vy*=-1; }
        draw() { ctx.beginPath(); ctx.arc(this.x,this.y,this.r,0,Math.PI*2); ctx.fillStyle=`rgba(99,102,241,${this.o})`; ctx.fill(); } }
    for(let i=0;i<Math.min(40,Math.floor(window.innerWidth/35));i++) particles.push(new P());
    (function animate() { ctx.clearRect(0,0,cvs.width,cvs.height);
        for(let i=0;i<particles.length;i++) for(let j=i+1;j<particles.length;j++) { const dx=particles[i].x-particles[j].x, dy=particles[i].y-particles[j].y, d=Math.sqrt(dx*dx+dy*dy);
            if(d<150){ctx.beginPath();ctx.moveTo(particles[i].x,particles[i].y);ctx.lineTo(particles[j].x,particles[j].y);ctx.strokeStyle=`rgba(99,102,241,${0.08*(1-d/150)})`;ctx.lineWidth=0.5;ctx.stroke();}}
        particles.forEach(p=>{p.update();p.draw();}); requestAnimationFrame(animate); })();
}

// ─── Scroll Animations ─────────────────────────────────────────────────────────
function initScrollAnimations() {
    const obs = new IntersectionObserver(entries => entries.forEach(e => { if(e.isIntersecting){e.target.classList.add('visible');obs.unobserve(e.target);} }), { threshold: 0.05 });
    document.querySelectorAll('.kpi-card,.chart-card,.leader-table,.data-table,.workflow-step').forEach(el => {
        el.style.opacity='0'; el.style.transform='translateY(15px)'; el.style.transition='opacity 0.4s ease, transform 0.4s ease'; obs.observe(el);
    });
}
const sty = document.createElement('style');
sty.textContent = '.kpi-card.visible,.chart-card.visible,.leader-table.visible,.data-table.visible,.workflow-step.visible{opacity:1!important;transform:translateY(0)!important;}';
document.head.appendChild(sty);

console.log('%c 🚀 AI Velocity Portal v3.0 %c Leaderboard Mode ', 'background:#6366f1;color:white;padding:8px 16px;border-radius:4px;font-weight:bold;', 'background:#10b981;color:white;padding:8px 12px;border-radius:4px;');
