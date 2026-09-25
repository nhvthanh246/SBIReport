"use strict";

const PAGES = [
  { key: "SBIR", name: "SBI Remit Việt Nam", url: "https://www.facebook.com/SBIVIETNAM/?locale=vi_VN" },
  { key: "SMILES", name: "Chuyển Tiền Smiles", url: "https://www.facebook.com/chuyentiensmiles/?locale=vi_VN" },
  { key: "DCOM", name: "DCOM Việt Nam", url: "https://www.facebook.com/DCOMvietnam/?locale=vi_VN" }
];

let running = false;

chrome.runtime.onInstalled.addListener(() => {
  chrome.storage.local.set({
    runState: { status: "idle", pages: [], message: "Sẵn sàng" }
  });
});

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === "GET_STATE") {
    chrome.storage.local.get("runState").then(({ runState }) => sendResponse({ ok: true, runState }));
    return true;
  }
  if (message?.type === "START_COLLECTION") {
    if (running) {
      sendResponse({ ok: false, error: "Một lượt thu thập đang chạy." });
      return false;
    }
    running = true;
    collectAll()
      .then((result) => sendResponse({ ok: true, result }))
      .catch((error) => sendResponse({ ok: false, error: error?.message || String(error) }))
      .finally(() => { running = false; });
    return true;
  }
  return false;
});

async function collectAll() {
  const results = [];
  await saveState("running", results, "Đang bắt đầu…");

  for (let index = 0; index < PAGES.length; index += 1) {
    const page = PAGES[index];
    await saveState("running", results, `Đang đọc ${page.key} (${index + 1}/3)…`);
    try {
      results.push(await collectPage(page));
    } catch (error) {
      results.push({ ...page, ok: false, error: error?.message || String(error) });
    }
    await saveState("running", results, `Đã đọc ${index + 1}/3 fanpage`);
  }

  const successCount = results.filter((item) => item.ok).length;
  const approximateCount = results.filter((item) => item.ok && !item.followers?.exact).length;
  const status = successCount === PAGES.length ? "complete" : "partial";
  const message = successCount === PAGES.length && approximateCount === 0
    ? "Đã lấy đủ ba số Follow chính xác. Có thể copy cột Excel."
    : successCount === PAGES.length
      ? `Đã lấy đủ Follow, nhưng ${approximateCount}/3 trang chỉ cung cấp số làm tròn.`
    : `Lấy được ${successCount}/3 trang. Xem chi tiết lỗi bên dưới.`;
  await saveState(status, results, message);
  return results;
}

async function collectPage(page) {
  const urls = pageUrls(page.url);
  const tab = await chrome.tabs.create({ url: "about:blank", active: false });
  let debuggerAttached = false;
  try {
    try {
      await enableMobileEmulation(tab.id);
      debuggerAttached = true;
    } catch (error) {
      console.warn("Không bật được giả lập mobile, tiếp tục với chế độ thường:", error);
    }

    const combined = { followers: null, attempts: [], mobileEmulation: debuggerAttached };
    let lastError = "Trang chưa trả về dữ liệu.";

    for (let urlIndex = 0; urlIndex < urls.length; urlIndex += 1) {
      await chrome.tabs.update(tab.id, { url: urls[urlIndex] });
      await waitForTabComplete(tab.id, 35000);
      await delay(3000);

      for (let attempt = 0; attempt < 2; attempt += 1) {
        try {
          const response = await extractFromTab(tab.id);
          combined.attempts.push({ url: response?.url || urls[urlIndex], diagnostics: response?.diagnostics });
          if (response?.followers && (!combined.followers || response.followers.exact)) combined.followers = response.followers;
          if (combined.followers?.exact) {
            return {
              ...page,
              ok: true,
              followers: combined.followers,
              attempts: combined.attempts,
              mobileEmulation: combined.mobileEmulation,
              url: response.url,
              title: response.title
            };
          }
          if (response?.error) lastError = response.error;
        } catch (error) {
          lastError = error?.message || String(error);
        }
        await delay(1800);
      }
    }
    if (combined.followers) {
      return {
        ...page,
        ok: true,
        followers: combined.followers,
        attempts: combined.attempts,
        mobileEmulation: combined.mobileEmulation,
        url: urls[urls.length - 1]
      };
    }
    throw new Error(`${lastError} Đã thử cả trang Giới thiệu và trang chính.`);
  } finally {
    if (debuggerAttached) {
      try { await detachDebugger(tab.id); } catch (_error) { /* tab có thể đã đóng */ }
    }
    try { await chrome.tabs.remove(tab.id); } catch (_error) { /* tab đã đóng */ }
  }
}

async function enableMobileEmulation(tabId) {
  const target = { tabId };
  await attachDebugger(target);
  try {
    await sendDebuggerCommand(target, "Network.enable", {});
    await sendDebuggerCommand(target, "Network.setUserAgentOverride", {
      userAgent: "Mozilla/5.0 (Linux; Android 14; Pixel 7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Mobile Safari/537.36",
      acceptLanguage: "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
      platform: "Android"
    });
    await sendDebuggerCommand(target, "Emulation.setDeviceMetricsOverride", {
      width: 390,
      height: 844,
      deviceScaleFactor: 3,
      mobile: true,
      screenWidth: 390,
      screenHeight: 844,
      positionX: 0,
      positionY: 0
    });
    await sendDebuggerCommand(target, "Emulation.setTouchEmulationEnabled", {
      enabled: true,
      maxTouchPoints: 5
    });
  } catch (error) {
    await detachDebugger(tabId).catch(() => {});
    throw error;
  }
}

function attachDebugger(target) {
  return new Promise((resolve, reject) => {
    chrome.debugger.attach(target, "1.3", () => {
      const error = chrome.runtime.lastError;
      if (error) reject(new Error(error.message));
      else resolve();
    });
  });
}

function sendDebuggerCommand(target, method, params) {
  return new Promise((resolve, reject) => {
    chrome.debugger.sendCommand(target, method, params, (result) => {
      const error = chrome.runtime.lastError;
      if (error) reject(new Error(`${method}: ${error.message}`));
      else resolve(result);
    });
  });
}

function detachDebugger(tabId) {
  return new Promise((resolve, reject) => {
    chrome.debugger.detach({ tabId }, () => {
      const error = chrome.runtime.lastError;
      if (error) reject(new Error(error.message));
      else resolve();
    });
  });
}

async function extractFromTab(tabId) {
  try {
    return await chrome.tabs.sendMessage(tabId, { type: "EXTRACT_STATS" });
  } catch (firstError) {
    if (!/Receiving end does not exist|Could not establish connection/i.test(firstError?.message || "")) throw firstError;
    await chrome.scripting.executeScript({ target: { tabId }, files: ["parser.js", "content.js"] });
    await delay(300);
    return chrome.tabs.sendMessage(tabId, { type: "EXTRACT_STATS" });
  }
}

function pageUrls(mainUrl) {
  const parsed = new URL(mainUrl);
  const basePath = parsed.pathname.replace(/\/+$/, "");
  const about = new URL(parsed.toString());
  about.pathname = `${basePath}/about/`;
  return [parsed.toString(), about.toString()];
}

async function waitForTabComplete(tabId, timeoutMs) {
  const current = await chrome.tabs.get(tabId);
  if (current.status === "complete") return;
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      reject(new Error("Facebook tải quá thời gian cho phép."));
    }, timeoutMs);
    function listener(updatedId, info) {
      if (updatedId !== tabId || info.status !== "complete") return;
      clearTimeout(timer);
      chrome.tabs.onUpdated.removeListener(listener);
      resolve();
    }
    chrome.tabs.onUpdated.addListener(listener);
  });
}

async function saveState(status, pages, message) {
  const runState = {
    status,
    pages,
    message,
    updatedAt: new Date().toISOString()
  };
  await chrome.storage.local.set({ runState });
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
