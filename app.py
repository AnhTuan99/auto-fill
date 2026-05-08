#!/usr/bin/env python3
"""
TNEX Survey Auto-Fill Tool — Web UI version
Run: python3 app.py  →  opens http://127.0.0.1:5000 in browser
"""

import json
import os
import queue
import threading
import time
import webbrowser
from datetime import datetime

import openpyxl
from flask import Flask, jsonify, render_template_string, request
from playwright.sync_api import sync_playwright

os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"
# ── Config ───────────────────────────────────────────────────────────────────

FORM_URL = (
    "https://forms.cloud.microsoft/Pages/ResponsePage.aspx"
    "?id=8tFN8x3T6E2GNbA2DFkW-SAWx3QxpEVKooBtoV3baItUQUgxSTlZWkhLRU5DVlpRN0xSWDQ0UjdDSC4u"
)
EXCEL_PATH = "40_unique_long_responses_soft_q4.xlsx"

app = Flask(__name__)
log_queue: queue.Queue = queue.Queue()
stop_event = threading.Event()
state = {"running": False, "done": 0, "total": 0, "success": 0, "fail": 0}

# ── Excel ─────────────────────────────────────────────────────────────────────


def load_data():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    return [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]


DATA = load_data()

# ── Form automation ──────────────────────────────────────────────────────────


def fill_single_form(page, row_data, row_num: int) -> bool:
    try:
        log_queue.put(f"[Hàng {row_num}] Đang tải form...")
        page.goto(FORM_URL, wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(2000)

        questions = page.query_selector_all('[data-automation-id="questionItem"]')
        if len(questions) < 6:
            log_queue.put(f"[Hàng {row_num}] ❌ Chỉ thấy {len(questions)} câu (cần 6)")
            return False

        # Q1 – radio Phòng ban
        q1_val = str(row_data[0]).strip() if row_data[0] else ""
        for r in questions[0].query_selector_all('input[type="radio"]'):
            if (r.get_attribute("value") or "").lower() == q1_val.lower():
                r.click(force=True)
                log_queue.put(f"[Hàng {row_num}] Q1 ✓  '{r.get_attribute('value')}'")
                break
        page.wait_for_timeout(400)

        # Q2 – radio Công ty cũ
        q2_val = str(row_data[1]).strip() if row_data[1] else ""
        for r in questions[1].query_selector_all('input[type="radio"]'):
            if (r.get_attribute("value") or "").lower() == q2_val.lower():
                r.click(force=True)
                log_queue.put(f"[Hàng {row_num}] Q2 ✓  '{r.get_attribute('value')}'")
                break
        page.wait_for_timeout(400)

        # Q3–Q6 – textarea
        for q_idx in range(2, 6):
            text = str(row_data[q_idx]).strip() if row_data[q_idx] else ""
            ta = questions[q_idx].query_selector("textarea")
            if ta and text:
                ta.click()
                ta.fill(text)
                log_queue.put(f"[Hàng {row_num}] Q{q_idx + 1} ✓  {len(text)} ký tự")
                page.wait_for_timeout(300)

        # Submit
        submit = page.query_selector('[data-automation-id="submitButton"]')
        if not submit:
            log_queue.put(f"[Hàng {row_num}] ❌ Không tìm thấy nút Submit")
            return False
        submit.click()
        page.wait_for_timeout(4000)
        log_queue.put(f"[Hàng {row_num}] ✅ Gửi thành công!")
        return True

    except Exception as e:
        log_queue.put(f"[Hàng {row_num}] ❌ Lỗi: {e}")
        return False


def run_automation(indices: list, delay: int, show_browser: bool):
    state.update(running=True, done=0, total=len(indices), success=0, fail=0)
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=not show_browser,
                args=["--no-sandbox", "--disable-dev-shm-usage"],
            )
            page = browser.new_context().new_page()

            for done, idx in enumerate(indices):
                if stop_event.is_set():
                    break
                log_queue.put(
                    f"── Hàng {idx + 1}  ({done + 1}/{len(indices)}) ────────────────"
                )
                ok = fill_single_form(page, DATA[idx], idx + 1)
                state["success"] += ok
                state["fail"] += not ok
                state["done"] = done + 1

                if done < len(indices) - 1 and not stop_event.is_set():
                    log_queue.put(f"⏳ Chờ {delay}s...")
                    for _ in range(delay * 10):
                        if stop_event.is_set():
                            break
                        time.sleep(0.1)

            browser.close()
    except Exception as e:
        log_queue.put(f"❌ Lỗi nghiêm trọng: {e}")
    finally:
        s, f = state["success"], state["fail"]
        log_queue.put(f"━━━ Hoàn thành: ✅ {s} thành công  ❌ {f} thất bại ━━━")
        state["running"] = False


# ── Flask routes ──────────────────────────────────────────────────────────────


@app.route("/")
def index():
    rows = []
    for i, r in enumerate(DATA):

        def trunc(v, n=50):
            s = str(v) if v else ""
            return s[:n] + "…" if len(s) > n else s

        rows.append({"index": i, "cols": [trunc(r[j]) for j in range(6)]})
    return render_template_string(HTML_TEMPLATE, rows=rows, total=len(DATA))


@app.route("/start", methods=["POST"])
def start():
    if state["running"]:
        return jsonify(error="Đang chạy"), 400
    body = request.json
    indices = body.get("indices", [])
    delay = int(body.get("delay", 3))
    show_browser = bool(body.get("show_browser", True))
    if not indices:
        return jsonify(error="Không có hàng nào được chọn"), 400
    stop_event.clear()
    threading.Thread(
        target=run_automation, args=(indices, delay, show_browser), daemon=True
    ).start()
    return jsonify(ok=True)


@app.route("/stop", methods=["POST"])
def stop():
    stop_event.set()
    return jsonify(ok=True)


@app.route("/status")
def status():
    logs = []
    while not log_queue.empty():
        logs.append(log_queue.get_nowait())
    return jsonify(**state, logs=logs)


# ── HTML/JS UI ────────────────────────────────────────────────────────────────

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>TNEX Auto-Fill Tool</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
         background: #f0f4f8; color: #1a1a2e; }
  header { background: #1565c0; color: white; padding: 14px 24px;
           display: flex; align-items: center; gap: 16px; }
  header h1 { font-size: 18px; font-weight: 700; }
  header span { font-size: 12px; color: #90caf9; }
  .main { padding: 16px 24px; display: flex; flex-direction: column; gap: 12px; }

  /* Control panel */
  .panel { background: white; border: 1px solid #e0e0e0; border-radius: 8px;
           padding: 14px 18px; }
  .panel h3 { font-size: 13px; color: #555; margin-bottom: 10px; }
  .ctrl-row { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 8px; }
  .ctrl-row:last-child { margin-bottom: 0; }
  label { font-size: 13px; }
  input[type=number] { width: 64px; padding: 4px 8px; border: 1px solid #ccc;
                       border-radius: 5px; font-size: 13px; }
  input[type=checkbox] { width: 15px; height: 15px; cursor: pointer; }
  .btn { padding: 6px 14px; border: none; border-radius: 5px; cursor: pointer;
         font-size: 13px; font-weight: 600; transition: opacity .15s; }
  .btn:hover { opacity: .85; }
  .btn-blue  { background: #1976d2; color: white; }
  .btn-gray  { background: #e3f2fd; color: #1565c0; }
  .btn-red   { background: #e53935; color: white; }
  .btn-green { background: #2e7d32; color: white; font-size: 14px; padding: 8px 22px; }
  .btn:disabled { opacity: .45; cursor: default; }
  .sel-count { font-size: 13px; font-weight: 700; color: #1565c0; margin-left: 4px; }
  .divider { color: #bbb; }

  /* Table */
  .table-wrap { overflow: auto; max-height: 340px; border: 1px solid #e0e0e0;
                border-radius: 6px; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; }
  th { background: #1565c0; color: white; padding: 8px 10px; text-align: left;
       position: sticky; top: 0; z-index: 1; white-space: nowrap; }
  td { padding: 7px 10px; border-bottom: 1px solid #f0f0f0; vertical-align: top;
       max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  tr:nth-child(even) td { background: #f8f9fa; }
  tr.selected td { background: #bbdefb !important; }
  tr.active-row td { background: #fff9c4 !important; }
  td.chk-cell { text-align: center; cursor: pointer; font-size: 16px; width: 36px; }
  td.num-cell { text-align: center; color: #888; width: 36px; }

  /* Action bar */
  .action-bar { display: flex; align-items: center; gap: 16px; }
  .progress-wrap { flex: 1; }
  .prog-label { font-size: 12px; color: #555; margin-bottom: 4px; }
  progress { width: 100%; height: 10px; border-radius: 5px; overflow: hidden; }
  progress::-webkit-progress-bar { background: #e0e0e0; }
  progress::-webkit-progress-value { background: #1976d2; }

  /* Log */
  .log-box { background: #0d1117; color: #58d68d; font-family: 'Courier New', monospace;
             font-size: 12px; padding: 12px; border-radius: 6px; height: 180px;
             overflow-y: auto; white-space: pre-wrap; }
</style>
</head>
<body>
<header>
  <h1>TNEX Survey Auto-Fill Tool</h1>
  <span>CÂU HỎI KHẢO SÁT NỘI BỘ TNF &nbsp;·&nbsp; {{ total }} hàng dữ liệu</span>
</header>

<div class="main">

  <!-- Controls -->
  <div class="panel">
    <h3>Tuỳ chọn</h3>
    <div class="ctrl-row">
      <label>Chọn hàng:</label>
      <button class="btn btn-gray" onclick="selectAll()">Chọn tất cả</button>
      <button class="btn btn-gray" onclick="deselectAll()">Bỏ chọn tất cả</button>
      <span class="divider">|</span>
      <label>Chọn</label>
      <input type="number" id="firstN" value="10" min="1" max="{{ total }}">
      <label>hàng đầu:</label>
      <button class="btn btn-gray" onclick="selectFirstN()">Áp dụng</button>
      <span class="divider">|</span>
      <span class="sel-count" id="selCount">0 hàng được chọn</span>
    </div>
    <div class="ctrl-row">
      <label>Chờ giữa mỗi lần gửi:</label>
      <input type="number" id="delay" value="3" min="1" max="60">
      <label>giây</label>
      <span class="divider">|</span>
      <input type="checkbox" id="showBrowser" checked>
      <label for="showBrowser">Hiện cửa sổ trình duyệt</label>
    </div>
  </div>

  <!-- Table -->
  <div class="panel" style="padding: 10px 14px;">
    <h3 style="margin-bottom:8px">Dữ liệu Excel &nbsp;<small style="color:#aaa;font-weight:400">(click cột ✓ để chọn/bỏ chọn)</small></h3>
    <div class="table-wrap">
      <table id="dataTable">
        <thead>
          <tr>
            <th>✓</th><th>#</th>
            <th>Câu 1 – Phòng ban</th>
            <th>Câu 2 – Công ty cũ</th>
            <th>Câu 3 – Giờ làm việc</th>
            <th>Câu 4 – Bảo vệ tài sản</th>
            <th>Câu 5 – Quy định 5S</th>
            <th>Câu 6 – Trang phục</th>
          </tr>
        </thead>
        <tbody>
          {% for row in rows %}
          <tr id="row-{{ row.index }}" data-idx="{{ row.index }}">
            <td class="chk-cell" onclick="toggleRow({{ row.index }})">☐</td>
            <td class="num-cell">{{ row.index + 1 }}</td>
            {% for col in row.cols %}
            <td title="{{ col }}">{{ col }}</td>
            {% endfor %}
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </div>

  <!-- Action bar -->
  <div class="panel">
    <div class="action-bar">
      <button class="btn btn-green" id="startBtn" onclick="startFill()">▶ Bắt đầu điền form</button>
      <button class="btn btn-red"   id="stopBtn"  onclick="stopFill()" disabled>⏹ Dừng</button>
      <div class="progress-wrap">
        <div class="prog-label" id="progLabel">Sẵn sàng</div>
        <progress id="progBar" value="0" max="100"></progress>
      </div>
    </div>
  </div>

  <!-- Log -->
  <div class="panel" style="padding: 12px 14px;">
    <h3 style="margin-bottom:8px">Log hoạt động</h3>
    <div class="log-box" id="logBox"></div>
  </div>

</div>

<script>
const selected = new Set();
let polling = false;

function toggleRow(idx) {
  const tr = document.getElementById('row-' + idx);
  const td = tr.querySelector('.chk-cell');
  if (selected.has(idx)) {
    selected.delete(idx);
    td.textContent = '☐';
    tr.classList.remove('selected');
  } else {
    selected.add(idx);
    td.textContent = '☑';
    tr.classList.add('selected');
  }
  updateCount();
}

function selectAll() {
  document.querySelectorAll('#dataTable tbody tr').forEach(tr => {
    const idx = parseInt(tr.dataset.idx);
    selected.add(idx);
    tr.querySelector('.chk-cell').textContent = '☑';
    tr.classList.add('selected');
  });
  updateCount();
}

function deselectAll() {
  selected.clear();
  document.querySelectorAll('#dataTable tbody tr').forEach(tr => {
    tr.querySelector('.chk-cell').textContent = '☐';
    tr.classList.remove('selected');
  });
  updateCount();
}

function selectFirstN() {
  deselectAll();
  const n = parseInt(document.getElementById('firstN').value) || 10;
  document.querySelectorAll('#dataTable tbody tr').forEach(tr => {
    const idx = parseInt(tr.dataset.idx);
    if (idx < n) {
      selected.add(idx);
      tr.querySelector('.chk-cell').textContent = '☑';
      tr.classList.add('selected');
    }
  });
  updateCount();
}

function updateCount() {
  document.getElementById('selCount').textContent = selected.size + ' hàng được chọn';
}

function appendLog(msg) {
  const box = document.getElementById('logBox');
  const ts = new Date().toLocaleTimeString('vi-VN');
  box.textContent += '[' + ts + '] ' + msg + '\\n';
  box.scrollTop = box.scrollHeight;
}

async function startFill() {
  if (selected.size === 0) { alert('Vui lòng chọn ít nhất một hàng!'); return; }
  if (!confirm('Bắt đầu điền ' + selected.size + ' form?')) return;
  const delay = parseInt(document.getElementById('delay').value) || 3;
  const showBrowser = document.getElementById('showBrowser').checked;
  const res = await fetch('/start', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ indices: [...selected], delay, show_browser: showBrowser })
  });
  if (!res.ok) { const d = await res.json(); alert(d.error); return; }
  document.getElementById('startBtn').disabled = true;
  document.getElementById('stopBtn').disabled = false;
  startPolling();
}

async function stopFill() {
  await fetch('/stop', { method: 'POST' });
}

function startPolling() {
  if (polling) return;
  polling = true;
  const interval = setInterval(async () => {
    const res = await fetch('/status');
    const d = await res.json();
    d.logs.forEach(appendLog);
    const prog = document.getElementById('progBar');
    prog.max = d.total || 1;
    prog.value = d.done;
    document.getElementById('progLabel').textContent =
      'Tiến độ: ' + d.done + '/' + d.total +
      '  |  ✅ ' + d.success + ' thành công' +
      '  |  ❌ ' + d.fail + ' thất bại';
    if (!d.running) {
      clearInterval(interval);
      polling = false;
      document.getElementById('startBtn').disabled = false;
      document.getElementById('stopBtn').disabled = true;
    }
  }, 500);
}
</script>
</body>
</html>
"""

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)
