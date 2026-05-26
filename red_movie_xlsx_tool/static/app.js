const $ = (id) => document.getElementById(id);

let defaultQuestions = [];

function setStatus(text) {
  $("status").textContent = text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

async function requestJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  let data = {};
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = { detail: text };
    }
  }
  if (!response.ok) {
    throw new Error(data.detail || `HTTP ${response.status}`);
  }
  return data;
}

function parseQuestions() {
  const questions = JSON.parse($("schemaEditor").value);
  if (!Array.isArray(questions)) {
    throw new Error("题目结构 JSON 必须是数组");
  }
  return questions;
}

function renderSchema(questions) {
  $("schemaEditor").value = JSON.stringify(questions, null, 2);
  $("questionCount").textContent = `${questions.length} 题`;
}

function renderResult(data) {
  const uniqueRows = data.summary?.unique_answer_rows ?? "";
  const separator = `${data.multi_separator || data.summary?.multi_separator || "┋"} ${data.multi_separator_codepoint || data.summary?.multi_separator_codepoint || "U+250B"}`;
  const distribution = data.summary
    ? `<details class="distribution">
        <summary>查看分布摘要</summary>
        <pre>${escapeHtml(JSON.stringify(data.summary, null, 2))}</pre>
      </details>`
    : "";

  $("result").innerHTML = `
    <div><strong>${escapeHtml(data.filename)}</strong></div>
    <div>样本：${escapeHtml(data.row_count)} 行，题目：${escapeHtml(data.question_count)} 题，唯一答案行：${escapeHtml(uniqueRows)}</div>
    <div>多选分隔符：<strong>${escapeHtml(separator)}</strong></div>
    <div class="path">${escapeHtml(data.path)}</div>
    <a class="download" href="${escapeHtml(data.download_url)}">下载 XLSX</a>
    ${distribution}
  `;
}

function renderFiles(files) {
  if (!files.length) {
    $("fileList").textContent = "暂无文件";
    return;
  }
  $("fileList").innerHTML = files
    .map(
      (file) => `
        <a href="${escapeHtml(file.download_url)}">
          <span>${escapeHtml(file.filename)}</span>
          <small>${Math.ceil(file.size / 1024)} KB</small>
        </a>
      `,
    )
    .join("");
}

async function loadSchema() {
  const data = await requestJson("/api/schema");
  defaultQuestions = data.questions;
  $("surveyTitle").textContent = data.title;
  $("separatorText").textContent = `${data.multi_separator} ${data.multi_separator_codepoint}`;
  renderSchema(defaultQuestions);
}

async function refreshFiles() {
  const data = await requestJson("/api/files");
  renderFiles(data.files || []);
}

async function generateWorkbook() {
  const count = Number($("count").value || 50);
  const seed = $("seed").value.trim() || null;
  const questions = parseQuestions();
  const data = await requestJson("/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ count, seed, questions }),
  });
  renderResult(data);
  await refreshFiles();
}

async function guarded(label, action) {
  try {
    setStatus(label);
    await action();
    setStatus("Ready");
  } catch (error) {
    setStatus("Error");
    $("result").innerHTML = `<span class="error">${escapeHtml(error.message || String(error))}</span>`;
  }
}

$("restoreSchema").addEventListener("click", () => {
  renderSchema(defaultQuestions);
});

$("refreshFiles").addEventListener("click", () => guarded("Refreshing", refreshFiles));
$("generateBtn").addEventListener("click", () => guarded("Generating", generateWorkbook));

guarded("Loading", async () => {
  await loadSchema();
  await refreshFiles();
});
