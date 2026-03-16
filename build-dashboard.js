const fs = require('fs');
const path = require('path');

// Collect all .tool.json files
const files = fs.readdirSync('.').filter(f => f.endsWith('.tool.json'));
const tools = files.map(f => JSON.parse(fs.readFileSync(f, 'utf8'))).sort((a, b) => a.order - b.order);

// Read tunnel config if exists
let tunnelConfig = {};
try {
  tunnelConfig = JSON.parse(fs.readFileSync('tunnel-config.json', 'utf8'));
} catch (e) {}

// Group by category
const categories = {
  'training-portal': { label: 'Training Portals', tools: [] },
  'online-tool': { label: 'Online Tools', tools: [] },
  'on-device': { label: 'On-Device Tools', tools: [] }
};

tools.forEach(t => {
  const cat = categories[t.category] || categories['online-tool'];
  cat.tools.push(t);
});

// Icon CSS class mapping
function iconClass(cat) {
  const map = {
    'training-portal': 'icon-sales',
    'online-tool': 'icon-banner',
    'on-device': 'icon-proposal'
  };
  return map[cat] || 'icon-banner';
}

// Build tool cards
function buildCard(tool) {
  const isServer = tool.type === 'server';
  const tagClass = isServer ? 'tag-local' : 'tag-online';
  const tagText = isServer ? 'Runs on Main Machine' : 'Available Online';
  const idAttr = isServer ? ` id="card-${tool.category}-${tool.order}"` : '';

  let actions = '';
  if (isServer) {
    const serverName = tool.name.toLowerCase().replace(/\s+/g, '-');
    actions = `
        <a class="btn btn-primary" href="#" target="_blank" data-server="${serverName}">Open</a>
        <button class="btn btn-secondary" onclick="startServer('${serverName}')">Start Server</button>
        <span class="status-dot checking" id="status-${serverName}"></span>
        <span class="status-text" id="status-text-${serverName}">Checking...</span>`;
  } else {
    actions = `<a class="btn btn-primary" href="${tool.file}" target="_blank">Open</a>`;
  }

  return `
    <div class="tool-card"${idAttr}>
      <div class="tool-icon ${iconClass(tool.category)}">${tool.icon}</div>
      <div class="tag ${tagClass}">${tagText}</div>
      <h3>${tool.name}</h3>
      <p>${tool.description}</p>
      <div class="tool-actions">${actions}
      </div>
    </div>`;
}

// Build server tools config for JS
const serverToolsJS = {};
tools.filter(t => t.type === 'server').forEach(t => {
  const name = t.name.toLowerCase().replace(/\s+/g, '-');
  serverToolsJS[name] = {
    port: t.server.port,
    localUrl: t.server.localUrl
  };
});

// Build sections HTML
let sectionsHTML = '';
for (const [key, cat] of Object.entries(categories)) {
  if (cat.tools.length === 0) continue;
  sectionsHTML += `
  <div class="section-label">${cat.label} <span class="count">${cat.tools.length}</span></div>
  <div class="tools-grid">
${cat.tools.map(t => buildCard(t)).join('\n')}
  </div>
`;
}

// Final HTML
const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dale Carnegie QLD - Tools Dashboard</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: Arial, Helvetica, sans-serif;
    background: #1a1a1a;
    color: #fff;
    min-height: 100vh;
  }
  header {
    background: #000;
    padding: 24px 40px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 3px solid #ffc708;
  }
  header h1 { font-size: 22px; font-weight: 600; letter-spacing: 0.5px; }
  header h1 span { color: #ffc708; }
  header .subtitle { color: #999; font-size: 13px; margin-top: 2px; }
  .env-badge {
    font-size: 11px; padding: 4px 12px; border-radius: 20px;
    font-weight: 600; text-transform: uppercase; letter-spacing: 1px;
  }
  .env-badge.online { background: #84c44c22; color: #84c44c; border: 1px solid #84c44c44; }
  .env-badge.local { background: #ffc70822; color: #ffc708; border: 1px solid #ffc70844; }
  .dashboard { max-width: 1100px; margin: 40px auto; padding: 0 24px; }
  .section-label {
    font-size: 12px; text-transform: uppercase; letter-spacing: 2px;
    color: #666; margin-bottom: 16px; display: flex; align-items: center; gap: 10px;
  }
  .section-label .count {
    background: #333; color: #999; padding: 2px 8px; border-radius: 10px; font-size: 11px;
  }
  .tools-grid {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
    gap: 20px; margin-bottom: 48px;
  }
  .tool-card {
    background: #242424; border: 1px solid #333; border-radius: 12px;
    padding: 28px; transition: border-color 0.2s, transform 0.15s;
    display: flex; flex-direction: column;
  }
  .tool-card:hover { border-color: #ffc708; transform: translateY(-2px); }
  .tool-icon {
    width: 48px; height: 48px; border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 22px; margin-bottom: 16px; flex-shrink: 0;
  }
  .tool-card h3 { font-size: 17px; font-weight: 600; margin-bottom: 8px; }
  .tool-card p { font-size: 13px; color: #999; line-height: 1.5; margin-bottom: 20px; flex-grow: 1; }
  .tool-actions { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
  .btn {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 9px 18px; border-radius: 8px; font-size: 13px;
    font-weight: 600; text-decoration: none; cursor: pointer;
    border: none; transition: opacity 0.15s;
  }
  .btn:hover { opacity: 0.85; }
  .btn-primary { background: #ffc708; color: #000; }
  .btn-secondary { background: #333; color: #ccc; }
  .status-dot { width: 8px; height: 8px; border-radius: 50%; display: inline-block; }
  .status-dot.online { background: #84c44c; }
  .status-dot.offline { background: #666; }
  .status-dot.checking { background: #ffc708; animation: pulse 1s infinite; }
  @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.3; } }
  .status-text { font-size: 12px; color: #666; margin-left: 4px; }
  .tool-card .tag {
    display: inline-block; font-size: 10px; text-transform: uppercase;
    letter-spacing: 1px; padding: 3px 8px; border-radius: 4px;
    background: #333; color: #999; margin-bottom: 12px; width: fit-content;
  }
  .tag.tag-online { background: #84c44c22; color: #84c44c; }
  .tag.tag-local { background: #ff634722; color: #ff6347; }
  .icon-sales { background: #0090cf22; color: #0090cf; }
  .icon-banner { background: #84c44c22; color: #84c44c; }
  .icon-proposal { background: #0090cf22; color: #0090cf; }
  .icon-feedback { background: #ffc70822; color: #ffc708; }
  footer { text-align: center; padding: 32px; color: #444; font-size: 12px; }
  .tool-card.add-new {
    border-style: dashed; border-color: #444;
    display: flex; align-items: center; justify-content: center;
    min-height: 220px; cursor: default;
  }
  .tool-card.add-new:hover { border-color: #666; transform: none; }
  .add-new-inner { text-align: center; color: #555; }
  .add-new-inner .plus { font-size: 36px; color: #444; margin-bottom: 8px; }
  .add-new-inner p { color: #555; font-size: 13px; }
</style>
</head>
<body>

<header>
  <div>
    <h1>Dale Carnegie <span>QLD</span> Tools</h1>
    <div class="subtitle">Internal tools dashboard</div>
  </div>
  <span class="env-badge" id="env-badge">Detecting...</span>
</header>

<div class="dashboard">
${sectionsHTML}
  <div class="tools-grid">
    <div class="tool-card add-new">
      <div class="add-new-inner">
        <div class="plus">+</div>
        <p>More tools coming soon</p>
      </div>
    </div>
  </div>
</div>

<footer>
  Dale Carnegie Queensland &mdash; William Farmer and Associates Pty Ltd
</footer>

<script>
  const isLocal = window.location.protocol === 'file:' || window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
  const badge = document.getElementById('env-badge');
  if (isLocal) {
    badge.textContent = 'Local';
    badge.className = 'env-badge local';
  } else {
    badge.textContent = 'Online';
    badge.className = 'env-badge online';
  }

  const serverTools = ${JSON.stringify(serverToolsJS, null, 2)};

  // Add tunnelUrl property
  Object.keys(serverTools).forEach(k => { serverTools[k].tunnelUrl = null; });

  async function loadTunnelConfig() {
    try {
      const resp = await fetch('tunnel-config.json?t=' + Date.now());
      if (resp.ok) {
        const config = await resp.json();
        Object.keys(config).forEach(k => {
          const name = k.toLowerCase().replace(/\\s+/g, '-');
          if (serverTools[name]) serverTools[name].tunnelUrl = config[k].url;
          // Also try exact key match
          if (serverTools[k]) serverTools[k].tunnelUrl = config[k].url;
        });
      }
    } catch (e) {}
  }

  function getToolUrl(name) {
    const tool = serverTools[name];
    if (!tool) return '#';
    if (isLocal) return tool.localUrl;
    return tool.tunnelUrl || tool.localUrl;
  }

  async function checkStatus(name) {
    const dot = document.getElementById('status-' + name);
    const text = document.getElementById('status-text-' + name);
    const openBtn = document.querySelector('[data-server="' + name + '"]');
    if (!dot || !text) return;

    const url = getToolUrl(name);
    if (openBtn) openBtn.href = url;

    dot.className = 'status-dot checking';
    text.textContent = 'Checking...';
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 4000);
      await fetch(url, { mode: 'no-cors', signal: controller.signal });
      clearTimeout(timeout);
      dot.className = 'status-dot online';
      text.textContent = 'Running';
    } catch (e) {
      dot.className = 'status-dot offline';
      if (!isLocal && !serverTools[name]?.tunnelUrl) {
        text.textContent = 'Tunnel not configured';
      } else {
        text.textContent = 'Offline';
      }
    }
  }

  function startServer(name) {
    if (!isLocal) {
      alert('The server must be started from the main machine.\\n\\nRemote control into the Mac and run:\\n~/proposal-generator/start-remote.sh');
      return;
    }
    const text = document.getElementById('status-text-' + name);
    if (text) text.textContent = 'Starting...';
    alert('To start: open Terminal and run the appropriate start script.');
    setTimeout(() => checkStatus(name), 3000);
  }

  loadTunnelConfig().then(() => {
    Object.keys(serverTools).forEach(name => checkStatus(name));
  });
  setInterval(() => {
    loadTunnelConfig().then(() => {
      Object.keys(serverTools).forEach(name => checkStatus(name));
    });
  }, 15000);
</script>

</body>
</html>`;

fs.writeFileSync('dashboard.html', html);
console.log(`Dashboard built with ${tools.length} tools from ${files.length} config files.`);
