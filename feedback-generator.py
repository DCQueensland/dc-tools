#!/usr/bin/env python3
"""
Dale Carnegie QLD - Feedback Form & QR Code Generator
Creates a Google Form (via Apps Script web app) and generates a QR code PNG.

SETUP: Deploy the Google Apps Script (see setup instructions in the web UI)
       and paste the web app URL below.
"""

import http.server
import json
import os
import webbrowser
import urllib.parse
import urllib.request
import qrcode
from pathlib import Path
from datetime import datetime

PORT = 8766
TOOLS_DIR = Path.home() / ".dc-tools"
OUTPUT_DIR = Path.home() / "Desktop"
CONFIG_FILE = TOOLS_DIR / "feedback-config.json"


def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE) as f:
            return json.load(f)
    return {}


def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=2)


HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dale Carnegie QLD - Feedback Form Generator</title>
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
    border-bottom: 3px solid #ffc708;
  }
  header .back { color: #ffc708; text-decoration: none; font-size: 13px; }
  header .back:hover { text-decoration: underline; }
  header h1 { font-size: 22px; font-weight: 600; margin-top: 4px; }
  header h1 span { color: #ffc708; }
  header .subtitle { color: #999; font-size: 13px; margin-top: 2px; }

  .container {
    max-width: 700px;
    margin: 40px auto;
    padding: 0 24px;
  }

  .card {
    background: #242424;
    border: 1px solid #333;
    border-radius: 12px;
    padding: 36px;
    margin-bottom: 24px;
  }

  .field { margin-bottom: 20px; }
  .field label {
    display: block;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: #888;
    margin-bottom: 8px;
  }
  .field input {
    width: 100%;
    padding: 12px 16px;
    background: #1a1a1a;
    border: 1px solid #444;
    border-radius: 8px;
    color: #fff;
    font-size: 15px;
    font-family: inherit;
  }
  .field input:focus { outline: none; border-color: #ffc708; }
  .field .hint { font-size: 11px; color: #666; margin-top: 6px; line-height: 1.4; }

  .btn-generate {
    display: block;
    width: 100%;
    padding: 14px;
    background: #ffc708;
    color: #000;
    border: none;
    border-radius: 8px;
    font-size: 16px;
    font-weight: 700;
    cursor: pointer;
    margin-top: 8px;
  }
  .btn-generate:hover { opacity: 0.85; }
  .btn-generate:disabled { opacity: 0.4; cursor: not-allowed; }

  .btn-save {
    display: inline-block;
    padding: 10px 20px;
    background: #333;
    color: #ccc;
    border: 1px solid #444;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    margin-top: 8px;
  }
  .btn-save:hover { border-color: #ffc708; color: #ffc708; }

  .status {
    margin-top: 24px;
    padding: 20px;
    border-radius: 8px;
    display: none;
    font-size: 14px;
    line-height: 1.5;
  }
  .status.loading { display: block; background: #2a2a1a; border: 1px solid #ffc708; color: #ffc708; }
  .status.success { display: block; background: #1a2a1a; border: 1px solid #84c44c; color: #84c44c; }
  .status.error   { display: block; background: #2a1a1a; border: 1px solid #e55; color: #e55; }

  .result { margin-top: 24px; display: none; }
  .result.show { display: block; }

  .result-row {
    display: flex;
    align-items: center;
    gap: 16px;
    background: #1a1a1a;
    border: 1px solid #333;
    border-radius: 8px;
    padding: 16px;
    margin-bottom: 12px;
  }
  .result-row .label { font-size: 11px; text-transform: uppercase; letter-spacing: 1px; color: #888; margin-bottom: 4px; }
  .result-row .value { font-size: 14px; word-break: break-all; }
  .result-row a { color: #ffc708; }
  .result-row .copy-btn {
    margin-left: auto;
    padding: 6px 14px;
    background: #333;
    border: 1px solid #444;
    border-radius: 6px;
    color: #ccc;
    font-size: 12px;
    cursor: pointer;
    flex-shrink: 0;
  }
  .result-row .copy-btn:hover { border-color: #ffc708; color: #ffc708; }

  .qr-preview { text-align: center; margin-top: 20px; }
  .qr-preview img { max-width: 250px; border-radius: 8px; border: 4px solid #333; }
  .qr-preview .qr-note { font-size: 12px; color: #666; margin-top: 8px; }

  .history { margin-top: 16px; }
  .history h3 { font-size: 14px; color: #666; margin-bottom: 12px; }
  .history-item {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #242424;
    border: 1px solid #333;
    border-radius: 8px;
    padding: 12px 16px;
    margin-bottom: 8px;
    font-size: 14px;
  }
  .history-item .name { color: #ccc; }
  .history-item .date { color: #666; font-size: 12px; }
  .history-item a { color: #ffc708; font-size: 12px; margin-left: 12px; }

  .setup-card {
    background: #2a2a1a;
    border: 1px solid #ffc708;
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 24px;
  }
  .setup-card h2 { font-size: 16px; color: #ffc708; margin-bottom: 16px; }
  .setup-card ol { padding-left: 20px; color: #ccc; font-size: 13px; line-height: 2; }
  .setup-card code {
    background: #1a1a1a;
    padding: 2px 8px;
    border-radius: 4px;
    font-size: 12px;
    color: #ffc708;
  }
  .setup-card .script-box {
    background: #1a1a1a;
    border: 1px solid #333;
    border-radius: 8px;
    padding: 16px;
    margin: 12px 0;
    font-family: monospace;
    font-size: 11px;
    color: #ccc;
    white-space: pre-wrap;
    word-break: break-all;
    max-height: 300px;
    overflow-y: auto;
    line-height: 1.5;
  }
  .setup-card .copy-script {
    display: inline-block;
    padding: 6px 16px;
    background: #ffc708;
    color: #000;
    border: none;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 700;
    cursor: pointer;
    margin-top: 4px;
  }

  .hidden { display: none; }
</style>
</head>
<body>

<header>
  <a href="dashboard.html" class="back">&larr; Back to Dashboard</a>
  <h1>Feedback Form <span>Generator</span></h1>
  <div class="subtitle">Creates a Google Form + QR code for any public workshop</div>
</header>

<div class="container">

  <!-- Setup panel (shown if no Apps Script URL configured) -->
  <div class="setup-card" id="setupPanel">
    <h2>One-Time Setup (2 minutes)</h2>
    <ol>
      <li>Go to <a href="https://script.google.com" target="_blank" style="color:#ffc708;">script.google.com</a> and click <strong>New Project</strong></li>
      <li>Delete everything in the editor and paste the script below</li>
      <li>Click <strong>Deploy &gt; New deployment</strong></li>
      <li>Select type: <strong>Web app</strong></li>
      <li>Set "Execute as" to <strong>Me</strong>, "Who has access" to <strong>Anyone</strong></li>
      <li>Click <strong>Deploy</strong>, then <strong>Authorize access</strong> (grant Forms permissions)</li>
      <li>Copy the <strong>Web app URL</strong> and paste it below</li>
    </ol>

    <div class="script-box" id="scriptContent">function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var title = data.title || "Workshop Feedback";
    var description = data.description || "";

    var form = FormApp.create(title);
    form.setDescription(description);
    form.setCollectEmail(false);
    form.setLimitOneResponsePerUser(false);

    form.addTextItem().setTitle("Full Name").setRequired(true);
    form.addTextItem().setTitle("Email Address").setRequired(true);
    form.addTextItem().setTitle("Best Contact Number").setRequired(true);
    form.addTextItem().setTitle("Company/Organisation Name").setRequired(true);
    form.addTextItem().setTitle("Title/Position").setRequired(true);

    form.addScaleItem()
      .setTitle("Please rate the workshop on a scale of 1 - 10")
      .setBounds(1, 10)
      .setLabels("Unsatisfactory", "Excellent")
      .setRequired(true);

    form.addParagraphTextItem()
      .setTitle("What was your biggest take away from today's workshop?")
      .setRequired(true);

    var interest = form.addMultipleChoiceItem()
      .setTitle("Would you be interested in learning how Dale Carnegie can support you in the future?")
      .setRequired(true);
    interest.setChoices([
      interest.createChoice("For Self"),
      interest.createChoice("For Team"),
      interest.createChoice("Both"),
      interest.createChoice("No thanks")
    ]);

    form.addParagraphTextItem()
      .setTitle("General Comments")
      .setRequired(false);

    var url = form.getPublishedUrl();
    var editUrl = form.getEditUrl();

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, formUrl: url, editUrl: editUrl })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok", message: "DC Feedback Form Creator is running" })
  ).setMimeType(ContentService.MimeType.JSON);
}</div>
    <button class="copy-script" onclick="copyScript()">Copy Script</button>

    <div class="field" style="margin-top: 20px;">
      <label>Apps Script Web App URL</label>
      <input type="url" id="appsScriptUrl" placeholder="https://script.google.com/macros/s/.../exec">
    </div>
    <button class="btn-save" onclick="saveSetup()">Save & Continue</button>
  </div>

  <!-- Main generator (shown after setup) -->
  <div class="card" id="mainPanel">
    <div class="field">
      <label>Workshop Name</label>
      <input type="text" id="workshopName" placeholder="e.g. Winning with Relationship Selling">
    </div>

    <div class="field">
      <label>Workshop Date</label>
      <input type="date" id="workshopDate" onclick="this.showPicker()" style="cursor: pointer;">
      <div class="hint">Click to select from calendar</div>
    </div>

    <button class="btn-generate" id="generateBtn" onclick="generate()">
      Generate Form & QR Code
    </button>

    <div class="status" id="status"></div>

    <div class="result" id="result">
      <div class="result-row">
        <div>
          <div class="label">Google Form Link (for respondents)</div>
          <div class="value"><a id="formLink" href="#" target="_blank"></a></div>
        </div>
        <button class="copy-btn" onclick="copyLink()">Copy</button>
      </div>

      <div class="result-row">
        <div>
          <div class="label">Edit Form (add/change questions)</div>
          <div class="value"><a id="editLink" href="#" target="_blank" style="color:#84c44c;">Open in Google Forms Editor</a></div>
        </div>
      </div>

      <div class="result-row">
        <div>
          <div class="label">QR Code Saved To</div>
          <div class="value" id="qrPath"></div>
        </div>
      </div>

      <div class="qr-preview">
        <img id="qrImage" src="" alt="QR Code">
        <div class="qr-note">QR code PNG saved to your Desktop — paste it into your visuals</div>
      </div>
    </div>
  </div>

  <div class="history" id="historySection">
    <h3>Recent Forms</h3>
    <div id="historyList"></div>
  </div>

  <div style="text-align:center; margin-top:24px;">
    <button class="btn-save" onclick="showSetup()" style="font-size:11px; color:#666; border-color:#333;">Change Apps Script URL</button>
  </div>
</div>

<script>
  let config = {};

  async function loadConfig() {
    try {
      const resp = await fetch('/config');
      config = await resp.json();
    } catch(e) { config = {}; }

    if (config.apps_script_url) {
      document.getElementById('setupPanel').classList.add('hidden');
      document.getElementById('mainPanel').classList.remove('hidden');
    } else {
      document.getElementById('setupPanel').classList.remove('hidden');
      document.getElementById('mainPanel').classList.add('hidden');
    }
  }

  function showSetup() {
    document.getElementById('setupPanel').classList.remove('hidden');
    if (config.apps_script_url) {
      document.getElementById('appsScriptUrl').value = config.apps_script_url;
    }
  }

  function copyScript() {
    const text = document.getElementById('scriptContent').textContent;
    navigator.clipboard.writeText(text);
    event.target.textContent = 'Copied!';
    setTimeout(() => event.target.textContent = 'Copy Script', 1500);
  }

  async function saveSetup() {
    const url = document.getElementById('appsScriptUrl').value.trim();
    if (!url || !url.startsWith('https://script.google.com/')) {
      alert('Please paste a valid Apps Script web app URL.');
      return;
    }
    try {
      await fetch('/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ apps_script_url: url })
      });
      config.apps_script_url = url;
      document.getElementById('setupPanel').classList.add('hidden');
      document.getElementById('mainPanel').classList.remove('hidden');
    } catch(e) {
      alert('Error saving: ' + e.message);
    }
  }

  async function generate() {
    const name = document.getElementById('workshopName').value.trim();
    const date = document.getElementById('workshopDate').value;

    if (!name) { alert('Enter a workshop name.'); return; }
    if (!date) { alert('Enter a workshop date.'); return; }

    const btn = document.getElementById('generateBtn');
    const status = document.getElementById('status');
    const result = document.getElementById('result');

    btn.disabled = true;
    btn.textContent = 'Generating...';
    result.classList.remove('show');
    status.className = 'status loading';
    status.textContent = 'Creating Google Form and generating QR code...';

    try {
      const resp = await fetch('/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ workshop_name: name, workshop_date: date })
      });

      const data = await resp.json();

      if (data.error) {
        status.className = 'status error';
        status.textContent = 'Error: ' + data.error;
        return;
      }

      status.className = 'status success';
      status.textContent = 'Done! Form created and QR code saved to Desktop.';

      document.getElementById('formLink').href = data.form_url;
      document.getElementById('formLink').textContent = data.form_url;
      document.getElementById('editLink').href = data.edit_url;
      document.getElementById('qrPath').textContent = data.qr_path;
      document.getElementById('qrImage').src = '/qr-image?path=' + encodeURIComponent(data.qr_path) + '&t=' + Date.now();

      result.classList.add('show');
      loadHistory();
    } catch (e) {
      status.className = 'status error';
      status.textContent = 'Error: ' + e.message;
    } finally {
      btn.disabled = false;
      btn.textContent = 'Generate Form & QR Code';
    }
  }

  function copyLink() {
    const url = document.getElementById('formLink').href;
    navigator.clipboard.writeText(url);
    event.target.textContent = 'Copied!';
    setTimeout(() => event.target.textContent = 'Copy', 1500);
  }

  async function loadHistory() {
    try {
      const resp = await fetch('/history');
      const data = await resp.json();
      const list = document.getElementById('historyList');
      if (!data.length) {
        list.innerHTML = '<div style="color:#555; font-size:13px;">No forms generated yet.</div>';
        return;
      }
      list.innerHTML = data.map(h => `
        <div class="history-item">
          <div>
            <div class="name">${h.name}</div>
            <div class="date">${h.date}</div>
          </div>
          <div>
            <a href="${h.form_url}" target="_blank">Form</a>
            <a href="${h.edit_url}" target="_blank">Edit</a>
          </div>
        </div>
      `).join('');
    } catch(e) {}
  }

  loadConfig();
  loadHistory();
</script>
</body>
</html>"""


class FeedbackHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def do_GET(self):
        if self.path == '/' or self.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html')
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode())

        elif self.path.startswith('/qr-image'):
            parsed = urllib.parse.urlparse(self.path)
            params = urllib.parse.parse_qs(parsed.query)
            qr_path = params.get('path', [''])[0]
            if qr_path and os.path.exists(qr_path):
                self.send_response(200)
                self.send_header('Content-Type', 'image/png')
                self.end_headers()
                with open(qr_path, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_response(404)
                self.end_headers()

        elif self.path == '/config':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(load_config()).encode())

        elif self.path == '/history':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps(load_history()).encode())

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        content_length = int(self.headers['Content-Length'])
        body = json.loads(self.rfile.read(content_length))

        if self.path == '/config':
            config = load_config()
            config.update(body)
            save_config(config)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': True}).encode())

        elif self.path == '/generate':
            try:
                result = create_form_and_qr(
                    body.get('workshop_name', '').strip(),
                    body.get('workshop_date', '').strip()
                )
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps(result).encode())
            except Exception as e:
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode())
        else:
            self.send_response(404)
            self.end_headers()


def format_date(date_str):
    d = datetime.strptime(date_str, '%Y-%m-%d')
    return d.strftime('%-d %B %Y')


def create_form_and_qr(workshop_name, workshop_date):
    config = load_config()
    apps_script_url = config.get('apps_script_url')
    if not apps_script_url:
        raise Exception("Apps Script URL not configured. Complete the setup first.")

    date_formatted = format_date(workshop_date)
    form_title = f"Dale Carnegie {workshop_name} Workshop Feedback"
    form_description = "Fill in the below fields to go into the draw to win a $1,000 credit to go towards any public Dale Carnegie Program! *Conditions Apply"

    # Call the Apps Script web app
    payload = json.dumps({
        'title': form_title,
        'description': form_description,
    }).encode()

    req = urllib.request.Request(
        apps_script_url,
        data=payload,
        headers={'Content-Type': 'application/json'},
        method='POST'
    )

    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode())

    if not result.get('success'):
        raise Exception(result.get('error', 'Unknown error from Apps Script'))

    form_url = result['formUrl']
    edit_url = result.get('editUrl', '')

    # Generate QR code PNG
    slug = workshop_name.lower().replace(' ', '-').replace('/', '-')
    slug = ''.join(c for c in slug if c.isalnum() or c == '-')
    qr_filename = f"dc-feedback-qr-{slug}-{workshop_date}.png"
    qr_path = str(OUTPUT_DIR / qr_filename)

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=20,
        border=2,
    )
    qr.add_data(form_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="#000000", back_color="#ffffff")
    img.save(qr_path)

    # Save to history
    save_to_history(workshop_name, date_formatted, form_url, edit_url, qr_path)

    return {
        'form_url': form_url,
        'edit_url': edit_url,
        'qr_path': qr_path,
    }


def load_history():
    history_file = TOOLS_DIR / "feedback-history.json"
    if history_file.exists():
        with open(history_file) as f:
            return json.load(f)
    return []


def save_to_history(name, date_formatted, form_url, edit_url, qr_path):
    history = load_history()
    history.insert(0, {
        'name': name,
        'date': date_formatted,
        'form_url': form_url,
        'edit_url': edit_url,
        'qr_path': qr_path,
        'created': datetime.now().isoformat(),
    })
    history = history[:50]
    history_file = TOOLS_DIR / "feedback-history.json"
    with open(history_file, 'w') as f:
        json.dump(history, f, indent=2)


if __name__ == '__main__':
    os.chdir(TOOLS_DIR)
    server = http.server.HTTPServer(('localhost', PORT), FeedbackHandler)
    print(f"Feedback Form Generator running at http://localhost:{PORT}")
    webbrowser.open(f"http://localhost:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        server.server_close()
