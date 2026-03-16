#!/usr/bin/env python3
"""
Dale Carnegie QLD - Eventbrite Banner Generator
Local web server that serves a form and triggers Claude Code to generate designs.
"""

import http.server
import json
import os
import subprocess
import webbrowser
import urllib.parse
from pathlib import Path

PORT = 8765
TOOLS_DIR = Path.home() / ".dc-tools"
REQUEST_FILE = TOOLS_DIR / "banner-request.json"

HTML_FORM = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dale Carnegie QLD - Banner Generator</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: Arial, Helvetica, sans-serif;
    background: #f5f5f5;
    color: #333;
    min-height: 100vh;
  }
  .header {
    background: #1a1a1a;
    padding: 20px 40px;
    display: flex;
    align-items: center;
    gap: 16px;
  }
  .header h1 { color: #fff; font-size: 20px; font-weight: bold; }
  .header .subtitle { color: #999; font-size: 13px; }
  .dc-mark {
    width: 40px; height: 40px; background: #ffc708; border-radius: 4px;
    display: flex; align-items: center; justify-content: center;
    font-weight: bold; font-size: 18px; color: #1a1a1a;
  }
  .container { max-width: 720px; margin: 40px auto; padding: 0 20px; }
  .card {
    background: #fff; border-radius: 12px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    padding: 36px; margin-bottom: 24px;
  }
  .card h2 {
    font-size: 18px; margin-bottom: 24px; color: #1a1a1a;
    border-bottom: 2px solid #ffc708; padding-bottom: 8px; display: inline-block;
  }
  .field { margin-bottom: 20px; }
  .field label {
    display: block; font-size: 13px; font-weight: bold; color: #555;
    margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.5px;
  }
  .field input, .field select, .field textarea {
    width: 100%; padding: 12px 14px; border: 2px solid #e0e0e0;
    border-radius: 8px; font-size: 15px; font-family: Arial, sans-serif;
    transition: border-color 0.2s; background: #fafafa;
  }
  .field input:focus, .field select:focus, .field textarea:focus {
    border-color: #0090cf; outline: none; background: #fff;
  }
  .row { display: flex; gap: 16px; }
  .row .field { flex: 1; }
  .radio-group { display: flex; gap: 12px; margin-top: 4px; }
  .radio-option { flex: 1; position: relative; }
  .radio-option input[type="radio"] { position: absolute; opacity: 0; width: 0; height: 0; }
  .radio-option label {
    display: flex; align-items: center; justify-content: center;
    padding: 12px; border: 2px solid #e0e0e0; border-radius: 8px;
    cursor: pointer; font-size: 14px; font-weight: bold;
    text-transform: none; letter-spacing: 0; color: #666;
    transition: all 0.2s; background: #fafafa;
  }
  .radio-option input[type="radio"]:checked + label {
    border-color: #0090cf; background: #e8f4fd; color: #0090cf;
  }
  .radio-option label:hover { border-color: #aaa; }
  .slider-container { display: flex; align-items: center; gap: 16px; }
  .slider-container input[type="range"] {
    flex: 1; height: 6px; -webkit-appearance: none;
    background: #e0e0e0; border-radius: 3px; border: none; padding: 0;
  }
  .slider-container input[type="range"]::-webkit-slider-thumb {
    -webkit-appearance: none; width: 24px; height: 24px;
    border-radius: 50%; background: #0090cf; cursor: pointer;
  }
  .slider-value {
    font-size: 24px; font-weight: bold; color: #0090cf;
    min-width: 30px; text-align: center;
  }
  .venue-input { display: none; margin-top: 10px; }
  .venue-input.active { display: block; }
  .btn-generate {
    width: 100%; padding: 16px; background: #0090cf; color: #fff;
    border: none; border-radius: 8px; font-size: 16px; font-weight: bold;
    cursor: pointer; transition: background 0.2s; letter-spacing: 0.5px;
  }
  .btn-generate:hover { background: #007ab8; }
  .btn-generate:active { background: #006a9e; }
  .btn-generate:disabled { background: #ccc; cursor: not-allowed; }
  .status-bar {
    display: none; margin-top: 24px; padding: 24px;
    border-radius: 12px; text-align: center;
  }
  .status-bar.working {
    display: block; background: #e8f4fd; border: 2px solid #0090cf;
  }
  .status-bar.success {
    display: block; background: #e8fbe8; border: 2px solid #84c44c;
  }
  .status-bar.error {
    display: block; background: #fde8e8; border: 2px solid #e44;
  }
  .status-bar h3 { margin-bottom: 8px; }
  .spinner {
    display: inline-block; width: 20px; height: 20px;
    border: 3px solid #0090cf33; border-top-color: #0090cf;
    border-radius: 50%; animation: spin 0.8s linear infinite;
    vertical-align: middle; margin-right: 8px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  .footer { text-align: center; padding: 30px; color: #aaa; font-size: 12px; }
</style>
</head>
<body>
<div class="header">
  <div class="dc-mark">DC</div>
  <div>
    <h1>Eventbrite Banner Generator</h1>
    <div class="subtitle">Dale Carnegie Queensland &mdash; Fill in details and click Generate</div>
  </div>
</div>
<div class="container">
  <div class="card">
    <h2>Workshop Details</h2>
    <div class="field">
      <label>Workshop Name</label>
      <input type="text" id="workshopName" placeholder="e.g. Leadership Training for Managers: Preview">
    </div>
    <div class="field">
      <label>Short Description (optional)</label>
      <input type="text" id="description" placeholder="e.g. Develop the skills to lead, inspire, and drive results.">
    </div>
    <div class="row">
      <div class="field">
        <label>Date</label>
        <input type="date" id="workshopDate">
      </div>
      <div class="field">
        <label>Start Time</label>
        <input type="time" id="startTime" value="12:00">
      </div>
      <div class="field">
        <label>End Time</label>
        <input type="time" id="endTime" value="13:00">
      </div>
    </div>
    <div class="row">
      <div class="field">
        <label>Timezone</label>
        <select id="timezone">
          <option value="AEST" selected>AEST (QLD)</option>
          <option value="AEDT">AEDT (NSW/VIC/TAS)</option>
          <option value="ACST">ACST (SA/NT)</option>
          <option value="AWST">AWST (WA)</option>
        </select>
      </div>
      <div class="field">
        <label>Price</label>
        <div class="radio-group">
          <div class="radio-option">
            <input type="radio" name="price" id="priceFree" value="free" checked>
            <label for="priceFree">FREE</label>
          </div>
          <div class="radio-option">
            <input type="radio" name="price" id="pricePaid" value="paid">
            <label for="pricePaid">Paid</label>
          </div>
        </div>
      </div>
    </div>
    <div class="field venue-input" id="priceField">
      <label>Price (incl. GST)</label>
      <input type="text" id="priceAmount" placeholder="e.g. $3,295 + GST">
    </div>
  </div>

  <div class="card">
    <h2>Venue / Platform</h2>
    <div class="field">
      <label>Format</label>
      <div class="radio-group">
        <div class="radio-option">
          <input type="radio" name="format" id="formatOnline" value="online" checked>
          <label for="formatOnline">Online</label>
        </div>
        <div class="radio-option">
          <input type="radio" name="format" id="formatInPerson" value="in-person">
          <label for="formatInPerson">In-Person</label>
        </div>
        <div class="radio-option">
          <input type="radio" name="format" id="formatHybrid" value="hybrid">
          <label for="formatHybrid">Hybrid</label>
        </div>
      </div>
    </div>
    <div id="onlinePlatform" class="field">
      <label>Platform</label>
      <select id="platform">
        <option value="Microsoft Teams">Microsoft Teams</option>
        <option value="Zoom">Zoom</option>
        <option value="Google Meet">Google Meet</option>
      </select>
    </div>
    <div id="venueField" class="field venue-input">
      <label>Venue Name & Location</label>
      <input type="text" id="venueName" placeholder="e.g. Mobo Co, South Brisbane">
    </div>
  </div>

  <div class="card">
    <h2>Design Options</h2>
    <div class="field">
      <label>Number of Design Options</label>
      <div class="slider-container">
        <input type="range" id="numDesigns" min="3" max="5" value="3">
        <span class="slider-value" id="numDesignsValue">3</span>
      </div>
    </div>
    <div class="field">
      <label>Banner Size</label>
      <select id="bannerSize">
        <option value="eventbrite" selected>Eventbrite (2160 x 1080)</option>
        <option value="email">Email Header (1200 x 400)</option>
        <option value="linkedin">LinkedIn Event (1584 x 396)</option>
        <option value="facebook">Facebook Event Cover (1920 x 1005)</option>
      </select>
    </div>
    <div class="field">
      <label>Colour Preference (optional)</label>
      <div class="radio-group">
        <div class="radio-option">
          <input type="radio" name="color" id="colorAny" value="any" checked>
          <label for="colorAny">Any</label>
        </div>
        <div class="radio-option">
          <input type="radio" name="color" id="colorBlue" value="blue">
          <label for="colorBlue" style="color:#0090cf">Blue</label>
        </div>
        <div class="radio-option">
          <input type="radio" name="color" id="colorGreen" value="green">
          <label for="colorGreen" style="color:#84c44c">Green</label>
        </div>
        <div class="radio-option">
          <input type="radio" name="color" id="colorYellow" value="yellow">
          <label for="colorYellow" style="color:#d4a800">Yellow</label>
        </div>
      </div>
    </div>
    <div class="field">
      <label>Additional Notes (optional)</label>
      <textarea id="notes" rows="3" placeholder="e.g. 'Use a professional photo background' or 'Make text larger'"></textarea>
    </div>
  </div>

  <button class="btn-generate" id="generateBtn" onclick="submitForm()">Generate Banners</button>

  <div class="status-bar" id="statusBar">
    <h3 id="statusTitle"></h3>
    <p id="statusMsg"></p>
  </div>
</div>
<div class="footer">Dale Carnegie Queensland &mdash; Banner Generator v1.0</div>

<script>
  document.querySelectorAll('input[name="format"]').forEach(radio => {
    radio.addEventListener('change', function() {
      const online = document.getElementById('onlinePlatform');
      const venue = document.getElementById('venueField');
      if (this.value === 'online') { online.style.display = 'block'; venue.classList.remove('active'); }
      else if (this.value === 'in-person') { online.style.display = 'none'; venue.classList.add('active'); }
      else { online.style.display = 'block'; venue.classList.add('active'); }
    });
  });
  document.querySelectorAll('input[name="price"]').forEach(radio => {
    radio.addEventListener('change', function() {
      document.getElementById('priceField').classList.toggle('active', this.value === 'paid');
    });
  });
  document.getElementById('numDesigns').addEventListener('input', function() {
    document.getElementById('numDesignsValue').textContent = this.value;
  });

  function formatDate(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr + 'T00:00:00');
    const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
    const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const day = d.getDate();
    const suffix = (day===1||day===21||day===31)?'st':(day===2||day===22)?'nd':(day===3||day===23)?'rd':'th';
    return days[d.getDay()] + ', ' + day + suffix + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
  }
  function formatTime(timeStr) {
    if (!timeStr) return '';
    const [h, m] = timeStr.split(':').map(Number);
    const suffix = h >= 12 ? 'pm' : 'am';
    const hour = h % 12 || 12;
    return m === 0 ? hour + ':00' + suffix : hour + ':' + m.toString().padStart(2,'0') + suffix;
  }

  function showStatus(type, title, msg) {
    const bar = document.getElementById('statusBar');
    bar.className = 'status-bar ' + type;
    document.getElementById('statusTitle').innerHTML = title;
    document.getElementById('statusMsg').innerHTML = msg;
  }

  async function submitForm() {
    const name = document.getElementById('workshopName').value.trim();
    const desc = document.getElementById('description').value.trim();
    const date = document.getElementById('workshopDate').value;
    const startTime = document.getElementById('startTime').value;
    const endTime = document.getElementById('endTime').value;
    const tz = document.getElementById('timezone').value;
    const format = document.querySelector('input[name="format"]:checked').value;
    const platform = document.getElementById('platform').value;
    const venue = document.getElementById('venueName').value.trim();
    const isFree = document.querySelector('input[name="price"]:checked').value === 'free';
    const price = document.getElementById('priceAmount').value.trim();
    const numDesigns = parseInt(document.getElementById('numDesigns').value);
    const bannerSize = document.getElementById('bannerSize').value;
    const colorPref = document.querySelector('input[name="color"]:checked').value;
    const notes = document.getElementById('notes').value.trim();

    if (!name) { alert('Please enter a workshop name.'); return; }
    if (!date) { alert('Please select a date.'); return; }

    let venueStr = '';
    if (format === 'online') venueStr = platform + ' (Online)';
    else if (format === 'in-person') venueStr = venue || 'TBC';
    else venueStr = platform + ' (Online) + ' + (venue || 'TBC');

    const data = {
      workshop_name: name,
      description: desc,
      date_raw: date,
      date_formatted: formatDate(date),
      start_time: formatTime(startTime),
      end_time: formatTime(endTime),
      timezone: tz,
      format: format,
      venue: venueStr,
      is_free: isFree,
      price: isFree ? 'FREE' : price,
      num_designs: numDesigns,
      banner_size: bannerSize,
      color_preference: colorPref,
      notes: notes
    };

    const btn = document.getElementById('generateBtn');
    btn.disabled = true;
    btn.textContent = 'Launching Claude...';
    showStatus('working', '<span class="spinner"></span> Launching Claude Code...', 'A new terminal window will open with Claude generating your banners.');

    try {
      const resp = await fetch('/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const result = await resp.json();
      if (result.success) {
        showStatus('success',
          'Claude Code is generating your banners!',
          'Switch to the <strong>Terminal window</strong> that just opened to watch progress and pick your favourite design.'
        );
        btn.textContent = 'Generate More Banners';
        btn.disabled = false;
      } else {
        throw new Error(result.error || 'Unknown error');
      }
    } catch (err) {
      showStatus('error', 'Something went wrong', err.message);
      btn.textContent = 'Generate Banners';
      btn.disabled = false;
    }
  }
</script>
</body>
</html>"""


class BannerHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # Suppress noisy logs

    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(HTML_FORM.encode())
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == "/generate":
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            # Save request to JSON file
            with open(REQUEST_FILE, "w") as f:
                json.dump(data, f, indent=2)

            # Build the prompt for Claude
            prompt = build_prompt(data)

            # Launch Claude Code in a new Terminal window
            try:
                launch_claude(prompt)
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": True}).encode())
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"success": False, "error": str(e)}).encode())
        else:
            self.send_response(404)
            self.end_headers()


def build_prompt(data):
    size_map = {
        "eventbrite": "Eventbrite banner (2160 x 1080 px)",
        "email": "Email header (1200 x 400 px)",
        "linkedin": "LinkedIn event cover (1584 x 396 px)",
        "facebook": "Facebook event cover (1920 x 1005 px)",
    }
    banner_size = size_map.get(data["banner_size"], data["banner_size"])
    n = data["num_designs"]

    prompt = f"""Create {n} banner designs in my Canva account for an upcoming workshop. Here are the details:

- Workshop: {data['workshop_name']}
- Date: {data['date_formatted']}
- Time: {data['start_time']} - {data['end_time']} {data['timezone']}
- Venue/Platform: {data['venue']}
- Price: {data['price']}
- Banner size: {banner_size}
- Include the Dale Carnegie logo: use black version (asset ID: MAHC9_OKQUc) on light backgrounds, white version (asset ID: MAHC9yRD74s) on dark backgrounds"""

    if data.get("description"):
        prompt += f"\n- Description: {data['description']}"
    if data.get("color_preference") and data["color_preference"] != "any":
        prompt += f"\n- Colour preference: {data['color_preference']} (from the DC brand palette)"
    if data.get("notes"):
        prompt += f"\n- Notes: {data['notes']}"

    prompt += f"""

Steps:
1. Generate {n} different design candidates using the DC brand kit (ID: kAGQmAKdOSw) as youtube_banner type
2. Save the best {n} as editable designs
3. Edit EACH design to ensure it shows: workshop title, date, time, venue/platform, price, and the Dale Carnegie Training logo
4. Show me all {n} previews so I can pick my favourite
5. After I choose, resize the winner to the correct banner dimensions"""

    return prompt


def launch_claude(prompt):
    """Launch Claude Code in a new Terminal window with the given prompt."""
    # Escape for AppleScript
    escaped = prompt.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
    applescript = f'''
    tell application "Terminal"
        activate
        do script "claude \\"{escaped}\\""
    end tell
    '''
    subprocess.run(["osascript", "-e", applescript], check=True)


def main():
    print("=" * 50)
    print("  Dale Carnegie QLD - Banner Generator")
    print("=" * 50)
    print(f"\n  Server running at: http://localhost:{PORT}")
    print("  Open the URL above in your browser.\n")
    print("  Press Ctrl+C to stop.\n")

    webbrowser.open(f"http://localhost:{PORT}")

    server = http.server.HTTPServer(("", PORT), BannerHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Server stopped.")
        server.server_close()


if __name__ == "__main__":
    main()
