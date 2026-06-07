// 단일 페이지 클라이언트 로직.
// 입력/출력 '내용'은 절대 렌더링하지 않는다. 파일명/종류/뱃지/건수/다운로드만.

const App = window.__APP__;
const API_BASE = (App.apiBase || "").replace(/\/$/, "");

function apiUrl(path) {
  return `${API_BASE}${path}`;
}

let sid = null;
let uploading = false;
const files = {}; // file_id -> {filename, ext, valid, reasons, suggested_slot}

// ── 설정 안내 배너 ───────────────────────────────────
function renderConfig() {
  const banner = document.getElementById("config-banner");
  const icon = document.getElementById("config-icon");
  const text = document.getElementById("config-text");
  if (App.hasKey) {
    banner.className = "mb-6 rounded-lg border border-emerald-200 bg-emerald-50 p-4 text-sm flex items-start gap-3";
    icon.textContent = "✅";
    text.innerHTML = `OpenAI 키가 설정되었습니다. (모델: <b>${App.model}</b>) 사업자명 정규화에 AI를 사용합니다.`;
  } else {
    banner.className = "mb-6 rounded-lg border border-amber-200 bg-amber-50 p-4 text-sm flex items-start gap-3";
    icon.textContent = "⚠️";
    text.innerHTML = `<b>secret.env</b> 에 <code>OPENAI_API_KEY</code> 를 넣어주세요. 키가 없으면 휴리스틱 폴백으로 정규화합니다. ` +
      `(<code>cp secret.env.example secret.env</code> 후 키 입력)`;
  }
}

// ── 세션 생성 ────────────────────────────────────────
async function ensureSession() {
  if (sid) return sid;
  const res = await fetch(apiUrl("/api/session"), { method: "POST" });
  const data = await res.json();
  sid = data.sid;
  return sid;
}

// ── 슬롯 select 생성 ─────────────────────────────────
function slotSelect(fileId, suggested) {
  const opts = App.slots.map(s =>
    `<option value="${s.key}" ${s.key === suggested ? "selected" : ""}>${s.label}</option>`
  ).join("");
  return `<select data-file="${fileId}" class="slot-select text-xs border rounded px-2 py-1 bg-white">
    <option value="">(미지정)</option>${opts}
  </select>`;
}

// ── 업로드 처리 중 표시 ───────────────────────────────
function setUploadProcessing(active, message) {
  const status = document.getElementById("upload-status");
  const text = document.getElementById("upload-status-text");
  const dz = document.getElementById("dropzone");
  const input = document.getElementById("file-input");
  if (active) {
    text.textContent = message;
    status.classList.remove("hidden");
    dz.classList.add("pointer-events-none", "opacity-60");
    input.disabled = true;
  } else {
    status.classList.add("hidden");
    dz.classList.remove("pointer-events-none", "opacity-60");
    input.disabled = false;
  }
}

// ── 파일 카드 렌더 ───────────────────────────────────
function renderFiles() {
  const list = document.getElementById("file-list");
  const hint = document.getElementById("empty-hint");
  const ids = Object.keys(files);
  hint.style.display = (uploading || ids.length) ? "none" : "block";
  list.innerHTML = ids.map(id => {
    const f = files[id];
    const badge = f.valid
      ? `<span class="badge badge-ok">검증 통과</span>`
      : `<span class="badge badge-fail">검증 실패</span>`;
    const reasons = (!f.valid && f.reasons && f.reasons.length)
      ? `<p class="text-xs text-red-600 mt-1">${f.reasons.join(" / ")}</p>` : "";
    const mapping = f.valid ? slotSelect(id, f.suggested_slot) : "";
    return `<div class="rounded-lg border bg-white p-3 flex items-center justify-between gap-3">
      <div class="min-w-0">
        <p class="text-sm font-medium truncate">${escapeHtml(f.filename)}</p>
        <div class="flex items-center gap-2 mt-1">${badge}<span class="text-xs text-slate-400">${f.ext || ""}</span></div>
        ${reasons}
      </div>
      <div class="shrink-0 flex items-center gap-2">
        ${mapping}
        <button type="button" data-del="${id}" title="삭제"
          class="del-btn flex h-7 w-7 items-center justify-center rounded-md border border-slate-200 text-slate-400 hover:border-red-300 hover:bg-red-50 hover:text-red-600">
          ✕
        </button>
      </div>
    </div>`;
  }).join("");

  list.querySelectorAll(".slot-select").forEach(sel => {
    sel.addEventListener("change", syncAssign);
  });
  list.querySelectorAll(".del-btn").forEach(btn => {
    btn.addEventListener("click", () => deleteFile(btn.dataset.del));
  });
}

// ── 파일 삭제 ────────────────────────────────────────
async function deleteFile(fileId) {
  if (!sid || !files[fileId]) return;
  try {
    const res = await fetch(apiUrl(`/api/file/${sid}/${fileId}`), { method: "DELETE" });
    if (!res.ok) throw new Error("삭제 실패");
    delete files[fileId];
    renderFiles();
    await syncAssign();
  } catch (e) {
    alert("파일 삭제에 실패했습니다.");
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function formatApiError(data, status) {
  if (!data) return `HTTP ${status}`;
  if (typeof data.detail === "string") return data.detail;
  if (Array.isArray(data.detail)) {
    return data.detail.map(d => d.msg || JSON.stringify(d)).join(", ");
  }
  return data.message || `HTTP ${status}`;
}

// ── 슬롯 매핑 동기화 ─────────────────────────────────
async function syncAssign() {
  const assign = {};
  document.querySelectorAll(".slot-select").forEach(sel => {
    const slot = sel.value;
    if (slot) assign[slot] = sel.dataset.file; // 마지막 선택이 슬롯을 차지(1:1)
  });
  await fetch(apiUrl("/api/assign"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sid, assign }),
  });
}

// ── 업로드 처리 ──────────────────────────────────────
async function uploadFiles(fileList) {
  if (!fileList || !fileList.length) return;
  if (uploading) return;

  uploading = true;
  const total = fileList.length;
  setUploadProcessing(true, `파일 ${total}개 인식됨 — 검증·분류 처리 중…`);
  renderFiles();

  try {
    await ensureSession();
    let done = 0;
    for (const file of fileList) {
      setUploadProcessing(
        true,
        `파일 ${total}개 중 ${done + 1}번째 처리 중… (${file.name})`,
      );
      const fd = new FormData();
      fd.append("sid", sid);
      fd.append("file", file);
      try {
        const res = await fetch(apiUrl("/api/upload"), { method: "POST", body: fd });
        let data;
        try {
          data = await res.json();
        } catch {
          const timeoutHint = res.status === 502 || res.status === 504
            ? " (요청 시간이 초과되었을 수 있습니다. 페이지를 새로고침 후 다시 시도하세요.)"
            : "";
          alert(`업로드 실패 (${file.name}): 서버 응답을 읽을 수 없습니다.${timeoutHint}`);
          continue;
        }
        if (!res.ok) {
          const detail = formatApiError(data, res.status);
          alert(`업로드 실패 (${file.name}): ${detail}`);
          continue;
        }
        if (!data.file_id) {
          alert(`업로드 실패 (${file.name}): file_id 가 없습니다.`);
          continue;
        }
        files[data.file_id] = data;
      } catch (e) {
        alert(`업로드 실패 (${file.name}): 네트워크 오류`);
      }
      done += 1;
    }
    renderFiles();
    await syncAssign();
  } finally {
    uploading = false;
    setUploadProcessing(false);
    renderFiles();
    document.getElementById("file-input").value = "";
  }
}

// ── 실행 ─────────────────────────────────────────────
async function run() {
  if (!sid) { alert("먼저 파일을 업로드하세요."); return; }
  await syncAssign();
  const btn = document.getElementById("run-btn");
  const prog = document.getElementById("progress");
  const result = document.getElementById("result");
  btn.disabled = true;
  prog.classList.remove("hidden");
  result.classList.add("hidden");

  try {
    const res = await fetch(apiUrl("/api/run"), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sid }),
    });
    const data = await res.json();
    renderResult(data);
  } catch (e) {
    renderResult({ ok: false, message: "처리 중 오류가 발생했습니다." });
  } finally {
    btn.disabled = false;
    prog.classList.add("hidden");
  }
}

function renderResult(data) {
  const result = document.getElementById("result");
  const box = document.getElementById("result-box");
  result.classList.remove("hidden");
  if (!data.ok && (!data.downloads || !data.downloads.length)) {
    box.className = "rounded-lg border border-red-200 bg-red-50 p-4";
    box.innerHTML = `<p class="text-sm font-medium text-red-700">${escapeHtml(data.message || "실패")}</p>`;
    return;
  }
  box.className = "rounded-lg border border-emerald-200 bg-emerald-50 p-4";
  const downloads = (data.downloads || []).map(d =>
    `<a href="${apiUrl(d.url)}" class="inline-flex items-center gap-2 rounded-md bg-white border px-3 py-2 text-sm font-medium text-indigo-700 hover:bg-indigo-50 mr-2 mb-2">
       ⬇ ${escapeHtml(d.name)}
     </a>`).join("");
  box.innerHTML = `
    <p class="text-sm font-medium ${data.ok ? "text-emerald-700" : "text-amber-700"}">${escapeHtml(data.message)}</p>
    <p class="text-xs text-slate-500 mt-1">검토/오류 건수: <b>${data.error_count ?? 0}</b> (상세는 오류목록 파일 참고)</p>
    <div class="mt-3">${downloads}</div>`;
}

// ── 이벤트 바인딩 ────────────────────────────────────
function bind() {
  const dz = document.getElementById("dropzone");
  const input = document.getElementById("file-input");
  dz.addEventListener("click", () => input.click());
  input.addEventListener("change", e => uploadFiles(e.target.files));
  ["dragenter", "dragover"].forEach(ev =>
    dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("dragover"); }));
  ["dragleave", "drop"].forEach(ev =>
    dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("dragover"); }));
  dz.addEventListener("drop", e => uploadFiles(e.dataTransfer.files));
  document.getElementById("run-btn").addEventListener("click", run);
}

renderConfig();
bind();
