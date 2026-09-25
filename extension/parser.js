(function (root) {
  "use strict";

  const LABELS = {
    followers: ["người theo dõi", "người đang theo dõi", "lượt theo dõi", "followers", "follower"],
    likes: ["lượt thích", "người thích điều này", "người thích trang này", "likes", "people like this", "like this"]
  };

  function cleanText(value) {
    return String(value || "")
      .normalize("NFKC")
      .replace(/[\u00a0\u202f]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function parseHumanNumber(input) {
    const cleaned = cleanText(input).toUpperCase();
    const match = cleaned.match(/([0-9][0-9\s.,]*?)\s*(NGHÌN|NGHIN|TRIỆU|TRIEU|TỶ|TY|TR|K|M|B|N)?(?:\s|$)/i);
    if (!match) return null;

    let numeric = match[1].replace(/\s/g, "");
    const suffix = (match[2] || "").toUpperCase();
    let multiplier = 1;
    if (["K", "N", "NGHÌN", "NGHIN"].includes(suffix)) multiplier = 1e3;
    if (["M", "TR", "TRIỆU", "TRIEU"].includes(suffix)) multiplier = 1e6;
    if (["B", "TỶ", "TY"].includes(suffix)) multiplier = 1e9;

    if (multiplier > 1) {
      const lastComma = numeric.lastIndexOf(",");
      const lastDot = numeric.lastIndexOf(".");
      const decimalAt = Math.max(lastComma, lastDot);
      if (decimalAt >= 0) {
        numeric = numeric.slice(0, decimalAt).replace(/[.,]/g, "") + "." + numeric.slice(decimalAt + 1).replace(/[.,]/g, "");
      }
    } else {
      numeric = numeric.replace(/[.,]/g, "");
    }

    const parsed = Number(numeric);
    return Number.isFinite(parsed) ? Math.round(parsed * multiplier) : null;
  }

  function visibleCandidates(text, kind) {
    const normalized = cleanText(text);
    const labels = LABELS[kind].map(escapeRegExp).join("|");
    const number = "((?:[0-9]{1,3}(?:[ .,'’][0-9]{3})+|[0-9]+(?:[.,][0-9]+)?)\\s*(?:K|M|B|N|Tr|nghìn|nghin|triệu|trieu|tỷ|ty)?)";
    const regex = new RegExp(number + "\\s*(?:" + labels + ")", "giu");
    const found = [];
    let match;
    while ((match = regex.exec(normalized)) !== null) {
      const count = parseHumanNumber(match[1]);
      if (count !== null && count >= 100) {
        const display = cleanText(match[1]);
        found.push({ count, display, index: match.index, source: "visible", exact: !isAbbreviated(display) });
      }
    }
    return found;
  }

  function jsonCandidates(source, kind) {
    const keys = kind === "followers"
      ? ["followers_count", "follower_count", "subscriber_count", "page_followers_count", "followersCount", "followerCount"]
      : ["global_likers_count", "page_likers_count", "likers_count", "likes_count", "fan_count", "likeCount", "fanCount"];
    const keyPattern = keys.join("|");
    const variants = [
      new RegExp('["\\\\](?:' + keyPattern + ')["\\\\]\\s*:\\s*(\\d{2,})', "gi"),
      new RegExp('(?:' + keyPattern + ')%22%3A(\\d{2,})', "gi")
    ];
    const found = [];
    for (const regex of variants) {
      let match;
      while ((match = regex.exec(source)) !== null) {
        const count = Number(match[1]);
        if (Number.isSafeInteger(count) && count >= 100) {
          found.push({ count, display: String(count), index: match.index, source: "page-data", exact: true });
        }
      }
    }
    return deduplicate(found);
  }

  function deduplicate(items) {
    const seen = new Set();
    return items.filter((item) => {
      const key = item.count + ":" + item.source;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function chooseBest(visible, exact) {
    let bestVisible = null;
    if (visible.length) {
      const hint = visible[0];
      const maxDelta = Math.max(1000, hint.count * 0.06);
      bestVisible = visible
        .filter((item) => Math.abs(item.count - hint.count) <= maxDelta)
        .sort((a, b) => Number(b.exact) - Number(a.exact) || significantDigits(b.display) - significantDigits(a.display) || a.index - b.index)[0] || hint;
    }
    if (visible.length && exact.length) {
      const visibleHint = bestVisible || visible[0];
      const maxDelta = Math.max(1000, visibleHint.count * 0.06);
      const close = exact
        .map((item) => ({ item, delta: Math.abs(item.count - visibleHint.count) }))
        .filter((entry) => entry.delta <= maxDelta)
        .sort((a, b) => a.delta - b.delta || a.item.index - b.item.index);
      if (close.length) {
        return {
          count: close[0].item.count,
          display: visibleHint.display,
          source: "page-data+visible",
          exact: true
        };
      }
    }
    if (exact.length) return exact.sort((a, b) => a.index - b.index)[0];
    if (bestVisible) return bestVisible;
    return null;
  }

  function isAbbreviated(value) {
    return /(?:K|M|B|N|Tr|nghìn|nghin|triệu|trieu|tỷ|ty)\s*$/i.test(cleanText(value));
  }

  function significantDigits(value) {
    return (String(value).match(/\d/g) || []).length;
  }

  function extract(bodyText, pageSource) {
    const text = cleanText(bodyText);
    const source = String(pageSource || "");
    const sourceText = readableSource(source);
    const likesVisibleBody = visibleCandidates(text, "likes");
    const followersVisibleBody = visibleCandidates(text, "followers");
    const likesVisibleSource = visibleCandidates(sourceText, "likes");
    const followersVisibleSource = visibleCandidates(sourceText, "followers");
    const likesVisible = deduplicate([...likesVisibleBody, ...likesVisibleSource]);
    const followersVisible = deduplicate([...followersVisibleBody, ...followersVisibleSource]);
    const likes = chooseBest(likesVisible, jsonCandidates(sourceText, "likes"));
    const followers = chooseBest(followersVisible, jsonCandidates(sourceText, "followers"));
    return {
      likes,
      followers,
      diagnostics: {
        likesVisible: likesVisible.length,
        followersVisible: followersVisible.length,
        bodyLength: text.length,
        sourceLength: source.length
      }
    };
  }

  function readableSource(source) {
    return String(source || "")
      .replace(/\\u([0-9a-f]{4})/gi, (_match, hex) => String.fromCharCode(parseInt(hex, 16)))
      .replace(/\\x3c/gi, "<")
      .replace(/\\x3e/gi, ">")
      .replace(/&quot;|&#34;/gi, '"')
      .replace(/&amp;/gi, "&")
      .replace(/\\["/]/g, (match) => match.slice(1));
  }

  function excelValue(stat, preferDisplay) {
    if (!stat) return "";
    if (!preferDisplay || !stat.display || /^\d+$/.test(stat.display)) return String(stat.count);
    return stat.display
      .replace(/\s+/g, "")
      .replace(/nghìn|nghin|N$/i, "k")
      .replace(/triệu|trieu|Tr$/i, "m")
      .toLowerCase();
  }

  function escapeRegExp(value) {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  root.FBStatsParser = { extract, parseHumanNumber, excelValue, isAbbreviated };
})(globalThis);
