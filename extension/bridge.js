/* Cầu nối giữa extension và web tool làm báo cáo.
 *
 * VÌ SAO CẦN: web tool chạy trong trình duyệt như một trang thường, nên KHÔNG
 * thể tự đọc số follow của Facebook — CORS chặn, Facebook lại render bằng JS và
 * cần phiên đăng nhập, còn `tabs`/`scripting`/`debugger` là quyền chỉ extension
 * mới có. Nên extension vẫn là bên đi lấy số; file này chỉ đưa số sang tool để
 * người làm khỏi phải copy-paste.
 *
 * Chạy trong ngữ cảnh tách biệt của trang tool: nói chuyện với trang bằng
 * postMessage, nói chuyện với extension bằng chrome.runtime.
 */
(function () {
  "use strict";

  const TAG = "fbstats";

  function post(payload) {
    window.postMessage(Object.assign({ __fbstats: true }, payload), "*");
  }

  /* Chỉ đưa sang trang đúng những gì tool cần. Giữ lại url của fanpage, nội dung
     trang hay chẩn đoán gỡ lỗi trong extension — trang không cần biết. */
  function slimPage(page) {
    if (!page) return null;
    const stat = page.followers;
    return {
      key: page.key,
      ok: !!page.ok,
      error: page.error || null,
      followers: stat && Number.isFinite(stat.count)
        ? { count: stat.count, exact: !!stat.exact, display: String(stat.display || "") }
        : null
    };
  }

  function slimState(state) {
    return {
      status: state?.status || "idle",
      message: state?.message || "",
      pages: Array.isArray(state?.pages) ? state.pages.map(slimPage) : []
    };
  }

  async function handle(request) {
    const type = request.action === "collect" ? "START_COLLECTION" : "GET_STATE";
    const response = await chrome.runtime.sendMessage({ type });
    if (type === "GET_STATE") return slimState(response?.runState);
    if (!response?.ok) throw new Error(response?.error || "Không chạy được lượt thu thập.");
    const after = await chrome.runtime.sendMessage({ type: "GET_STATE" });
    return slimState(after?.runState);
  }

  window.addEventListener("message", (event) => {
    if (event.source !== window) return;
    const data = event.data;
    if (!data || data.__fbstats !== true || data.kind !== "request") return;
    handle(data)
      .then((state) => post({ kind: "response", id: data.id, ok: true, state }))
      .catch((error) => post({ kind: "response", id: data.id, ok: false, error: error?.message || String(error) }));
  });

  /* Tiến độ: background ghi runState vào storage sau mỗi fanpage, chuyển tiếp để
     tool hiện "Đang đọc SMILES (2/3)…" thay vì đứng im cả phút. */
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local" || !changes.runState?.newValue) return;
    post({ kind: "progress", state: slimState(changes.runState.newValue) });
  });

  post({ kind: "hello", version: chrome.runtime.getManifest().version, tag: TAG });
})();
