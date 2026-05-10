"""Fix app.js: remove Owner/Submitter columns, add collapsible sections."""
import re

with open('frontend/app.js', 'r', encoding='utf-8') as f:
    content = f.read()

# Show what patterns exist
for pattern in ['<th>Owner</th>', '<th>Submitter</th>', 'ownerName(w)', 'ownerName(it)', "${it.name || 'Anonymous'}"]:
    idx = content.find(pattern)
    if idx >= 0:
        ctx = content[max(0,idx-30):idx+len(pattern)+30]
        print(f"Found '{pattern}' at pos {idx}: ...{repr(ctx)}...")
    else:
        print(f"NOT FOUND: '{pattern}'")

# 1. Remove <th>Owner</th> from completed table headers
content = content.replace('<th>Owner</th><th>Process</th><th>Leader</th><th>Impact</th><th>Stage</th><th>Replicable</th>',
                          '<th>Process</th><th>Leader</th><th>Impact</th><th>Stage</th><th>Replicable</th>')

# 2. Remove ownerName(w) td from completed table rows
content = re.sub(r'<td>\$\{ownerName\(w\)\}</td>\s*\n\s*<td class="col-process">',
                 '<td class="col-process">', content)

# 3. Remove <th>Submitter</th> everywhere
content = content.replace('<th>Submitter</th><th>Process</th><th>Leader</th><th>Status</th>',
                          '<th>Process</th><th>Leader</th><th>Status</th>')
content = content.replace('<th>Submitter</th><th>Process</th><th>Leader</th><th>Problem Statement</th>',
                          '<th>Process</th><th>Leader</th><th>Problem Statement</th>')
content = content.replace('<th>Submitter</th><th>Process</th><th>Type</th>',
                          '<th>Process</th><th>Type</th>')

# 4. Remove ${it.name || 'Anonymous'} td rows
content = re.sub(r"<td>\$\{it\.name \|\| 'Anonymous'\}</td>\s*\n\s*<td class=\"col-process\">",
                 '<td class="col-process">', content)

# 5. Remove ${ownerName(it)} td from duplicates
content = re.sub(r'<td>\$\{ownerName\(it\)\}</td>\s*\n\s*<td class="col-process">',
                 '<td class="col-process">', content)

# 6. Add collapsible section functionality
# Add CSS for collapsible at the end of the init function
collapsible_code = '''
// ─── Collapsible Sections ──────────────────────────────────────────────────────
function initCollapsible() {
    document.querySelectorAll('.subsection-title, .section-header .section-title').forEach(el => {
        // Don't make the main dashboard KPI section collapsible
        const section = el.closest('.section, .subsection');
        if (!section || section.id === 'dashboard') return;
        
        el.style.cursor = 'pointer';
        el.style.userSelect = 'none';
        
        // Add toggle arrow
        const arrow = document.createElement('span');
        arrow.className = 'collapse-arrow';
        arrow.textContent = ' ▾';
        arrow.style.cssText = 'transition:transform 0.3s;display:inline-block;margin-left:8px;font-size:14px;color:var(--text-muted);';
        el.appendChild(arrow);
        
        el.addEventListener('click', () => {
            const content = el.closest('.subsection') ? 
                el.nextElementSibling : 
                el.closest('.section-header')?.nextElementSibling;
            if (!content) return;
            
            const isHidden = content.style.display === 'none';
            content.style.display = isHidden ? '' : 'none';
            arrow.style.transform = isHidden ? '' : 'rotate(-90deg)';
            
            // Also hide subsequent siblings until next section-header
            if (!el.closest('.subsection')) {
                let sib = content.nextElementSibling;
                while (sib && !sib.classList.contains('section-header')) {
                    sib.style.display = isHidden ? '' : 'none';
                    sib = sib.nextElementSibling;
                }
            }
        });
    });
}
'''

# Insert collapsible init before the DOMContentLoaded or at end
if 'initCollapsible' not in content:
    content += collapsible_code
    # Add initCollapsible() call after data loads
    content = content.replace('initScrollAnimations();', 'initScrollAnimations();\n    initCollapsible();')

with open('frontend/app.js', 'w', encoding='utf-8') as f:
    f.write(content)

# Verify
c2 = open('frontend/app.js', encoding='utf-8').read()
print("\n--- VERIFICATION ---")
print(f"Owner col: {'<th>Owner</th>' in c2}")
print(f"Submitter col: {'<th>Submitter</th>' in c2}")
print(f"ownerName(w): {'ownerName(w)' in c2}")
print(f"ownerName(it): {'ownerName(it)' in c2}")
anon_check = "it.name || 'Anonymous'" in c2
print(f"it.name || Anonymous: {anon_check}")
print(f"initCollapsible: {'initCollapsible' in c2}")
print(f"Lines: {len(c2.splitlines())}")
