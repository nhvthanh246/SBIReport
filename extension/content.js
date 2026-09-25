(function () {
  "use strict";

  function collectAccessibleText() {
    const parts = [document.body?.innerText || ""];
    for (const element of document.querySelectorAll('meta[name="description"], meta[property="og:description"], [aria-label], [title]')) {
      const value = element.content || element.getAttribute("aria-label") || element.title;
      if (value) parts.push(value);
    }
    return parts.join("\n");
  }

  function collectPageSource() {
    return document.documentElement?.innerHTML || "";
  }

  function followerTargets() {
    return Array.from(document.querySelectorAll("a, span, div, [role='button']"))
      .filter((element) => /(?:người|lượt).*theo dõi|followers?/i.test(element.innerText || element.getAttribute("aria-label") || ""))
      .filter((element) => (element.innerText || "").length < 120)
      .filter((element) => element.getClientRects().length > 0)
      .sort((a, b) => (a.innerText || "").length - (b.innerText || "").length)
      .slice(0, 3);
  }

  async function revealFollowerDetails(target) {
    const clickable = target.closest("a, [role='button'], [tabindex]") || target;
    clickable.dispatchEvent(new MouseEvent("mouseover", { bubbles: true, view: window }));
    clickable.dispatchEvent(new MouseEvent("mouseenter", { bubbles: false, view: window }));
    await new Promise((resolve) => setTimeout(resolve, 250));
    if (globalThis.PointerEvent) {
      clickable.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, pointerType: "touch", isPrimary: true }));
      clickable.dispatchEvent(new PointerEvent("pointerup", { bubbles: true, pointerType: "touch", isPrimary: true }));
    }
    clickable.click();
    await new Promise((resolve) => setTimeout(resolve, 1500));
  }

  function extractDialogStats() {
    const dialogs = Array.from(document.querySelectorAll('[role="dialog"], [aria-modal="true"]'));
    if (!dialogs.length) return null;
    const text = dialogs.map((dialog) => dialog.innerText || dialog.getAttribute("aria-label") || "").join("\n");
    const source = dialogs.map((dialog) => dialog.outerHTML || "").join("\n");
    return FBStatsParser.extract(text, source);
  }

  async function extractStats() {
    let result = FBStatsParser.extract(collectAccessibleText(), collectPageSource());
    let hoverTargets = 0;
    if (result.followers && !result.followers.exact) {
      const targets = followerTargets();
      hoverTargets = targets.length;
      for (const target of targets) {
        await revealFollowerDetails(target);
        const dialogResult = extractDialogStats();
        if (dialogResult?.followers?.exact) {
          dialogResult.diagnostics.hoverTargets = hoverTargets;
          dialogResult.diagnostics.mobileDetailDialog = true;
          return dialogResult;
        }
      }
      result = FBStatsParser.extract(collectAccessibleText(), collectPageSource());
    }
    result.diagnostics.hoverTargets = hoverTargets;
    result.diagnostics.mobileDetailDialog = false;
    return result;
  }

  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    if (message?.type !== "EXTRACT_STATS") return false;
    extractStats().then((result) => {
      const missing = [!result.followers && "Follow"].filter(Boolean);
      sendResponse({
        ok: Boolean(result.followers),
        error: missing.length ? `Không tìm thấy ${missing.join(" và ")} trong trang Facebook đã tải.` : "",
        url: location.href,
        title: document.title,
        followers: result.followers,
        diagnostics: result.diagnostics
      });
    }).catch((error) => {
      sendResponse({ ok: false, error: error?.message || String(error), url: location.href });
    });
    return true;
  });
})();
