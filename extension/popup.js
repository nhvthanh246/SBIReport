"use strict";

const ORDER = ["SBIR", "SMILES", "DCOM"];
const LABELS = { SBIR: "SBIR", SMILES: "SMILES", DCOM: "DCOM" };
const elements = {
  collectButton: document.querySelector("#collectButton"),
  copyButton: document.querySelector("#copyButton"),
  statusBox: document.querySelector("#statusBox"),
  statusText: document.querySelector("#statusText"),
  resultRows: document.querySelector("#resultRows"),
  errorDetails: document.querySelector("#errorDetails"),
  columnPreview: document.querySelector("#columnPreview"),
  dateBadge: document.querySelector("#dateBadge")
};

let currentColumn = [];

elements.dateBadge.textContent = new Intl.DateTimeFormat("vi-VN", {
  timeZone: "Asia/Ho_Chi_Minh",
  day: "2-digit",
  month: "2-digit",
  year: "numeric"
}).format(new Date());

elements.collectButton.addEventListener("click", async () => {
  setRunning(true);
  try {
    const response = await chrome.runtime.sendMessage({ type: "START_COLLECTION" });
    if (!response?.ok) throw new Error(response?.error || "Không thể bắt đầu.");
  } catch (error) {
    renderState({ status: "partial", pages: [], message: error?.message || String(error) });
  } finally {
    setRunning(false);
    await loadState();
  }
});

elements.copyButton.addEventListener("click", async () => {
  if (currentColumn.length !== 6) return;
  try {
    await copyExcelColumn(currentColumn);
    const oldText = elements.copyButton.textContent;
    elements.copyButton.textContent = "Đã copy ✓";
    setTimeout(() => { elements.copyButton.textContent = oldText; }, 1600);
  } catch (_error) {
    elements.statusText.textContent = "Không copy được. Hãy chọn nội dung trong khung và nhấn Ctrl+C.";
  }
});

async function copyExcelColumn(values) {
  const plainText = values.join("\n");
  if (navigator.clipboard.write && globalThis.ClipboardItem) {
    const rows = values.map((value) => `<tr><td>${escapeHtml(value) || "<br>"}</td></tr>`).join("");
    const html = `<table>${rows}</table>`;
    await navigator.clipboard.write([new ClipboardItem({
      "text/plain": new Blob([plainText], { type: "text/plain" }),
      "text/html": new Blob([html], { type: "text/html" })
    })]);
    return;
  }
  await navigator.clipboard.writeText(plainText);
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[character]);
}

chrome.storage.onChanged.addListener((changes, area) => {
  if (area === "local" && changes.runState?.newValue) renderState(changes.runState.newValue);
});

async function loadState() {
  const response = await chrome.runtime.sendMessage({ type: "GET_STATE" });
  if (response?.runState) renderState(response.runState);
}

function renderState(state) {
  const pages = Array.isArray(state?.pages) ? state.pages : [];
  const byKey = Object.fromEntries(pages.map((page) => [page.key, page]));
  elements.statusBox.className = `status ${state?.status || "idle"}`;
  elements.statusText.textContent = state?.message || "Sẵn sàng";
  elements.resultRows.replaceChildren();
  elements.errorDetails.replaceChildren();
  const errors = [];

  for (const key of ORDER) {
    const page = byKey[key];
    const row = document.createElement("div");
    row.className = "result-row";
    const name = document.createElement("span");
    name.className = "page-name";
    name.textContent = LABELS[key];
    const likes = document.createElement("span");
    const followers = document.createElement("span");
    if (page?.ok) {
      likes.className = followers.className = "value";
      likes.textContent = "Bỏ qua";
      followers.textContent = formatStat(page.followers);
      if (!page.followers?.exact) followers.title = `Facebook chỉ hiển thị ${page.followers?.display || "số làm tròn"}`;
    } else if (page?.error) {
      likes.className = followers.className = "error";
      likes.textContent = "Lỗi";
      followers.textContent = "—";
      row.title = page.error;
      errors.push(`${key}: ${page.error}`);
    } else {
      likes.textContent = followers.textContent = "—";
    }
    row.append(name, likes, followers);
    elements.resultRows.append(row);
  }

  elements.errorDetails.hidden = errors.length === 0;
  for (const error of errors) {
    const line = document.createElement("p");
    line.textContent = error;
    elements.errorDetails.append(line);
  }

  currentColumn = buildColumn(byKey);
  const complete = currentColumn.length === 6 && [1, 3, 5].every((index) => Boolean(currentColumn[index]));
  elements.copyButton.disabled = !complete;
  elements.columnPreview.textContent = complete
    ? currentColumn.map((value) => value || "〈ô Like để trống〉").join("\n")
    : "〈ô Like để trống〉\n—\n〈ô Like để trống〉\n—\n〈ô Like để trống〉\n—";
  setRunning(state?.status === "running");
}

function buildColumn(byKey) {
  const values = [];
  for (const key of ORDER) {
    const page = byKey[key];
    if (!page?.ok) return [];
    values.push("");
    values.push(FBStatsParser.excelValue(page.followers, false));
  }
  return values;
}

function formatNumber(value) {
  return Number.isFinite(value) ? new Intl.NumberFormat("vi-VN").format(value) : "—";
}

function formatStat(stat) {
  if (!stat || !Number.isFinite(stat.count)) return "—";
  return `${stat.exact ? "" : "≈ "}${formatNumber(stat.count)}`;
}

function setRunning(isRunning) {
  elements.collectButton.disabled = isRunning;
  elements.collectButton.classList.toggle("running", isRunning);
  elements.collectButton.lastChild.textContent = isRunning ? " Đang lấy dữ liệu…" : " Lấy số liệu hôm nay";
}

loadState().catch(() => renderState({ status: "idle", pages: [], message: "Sẵn sàng" }));
