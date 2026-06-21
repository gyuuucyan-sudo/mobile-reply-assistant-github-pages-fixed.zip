const DB_URL = "https://keyboard-warrior-db-1443038288.cos.ap-shanghai.myqcloud.com/%E7%BB%88%E6%9E%81%E7%89%88_%E6%95%B0%E6%8D%AE%E5%BA%93%E4%BF%AE%E6%94%B920260620.before-period-break-20260620-133222.xlsx";
const state = { listings: [], rooms: [], rules: [], templates: [], listing: null, room: null, result: null };

const $ = (selector) => document.querySelector(selector);
const normalize = (text) => String(text || "").normalize("NFKC").toLowerCase().replace(/\s+/g, "");

init();

async function init() {
  bind();
  await loadDatabase();
}

function bind() {
  $("#listing").addEventListener("change", () => {
    state.listing = state.listings.find((item) => item["房源ID"] === $("#listing").value) || state.listings[0];
    fillRooms();
    showRoomOverview();
  });
  $("#room").addEventListener("change", () => {
    state.room = currentRooms().find((item) => item["房间号"] === $("#room").value) || currentRooms()[0];
    showRoomOverview();
  });
  $("#query").addEventListener("input", debounce(() => search($("#query").value), 120));
  $("#copy-ja").addEventListener("click", () => copy($("#ja").value));
  $("#copy-zh").addEventListener("click", () => copy($("#zh").value));
  $("#copy-en").addEventListener("click", () => copy($("#en").value));
}

async function loadDatabase() {
  try {
    $("#status").textContent = "正在读取腾讯云数据库...";
    const response = await fetchWithTimeout(`${DB_URL}?_=${Date.now()}`, { cache: "no-store" }, 12000);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const workbook = XLSX.read(await response.arrayBuffer(), { type: "array" });
    const sheet = (name) => XLSX.utils.sheet_to_json(workbook.Sheets[name] || {}, { defval: "" });
    applyDatabase({
      listings: sheet("房源信息"),
      rooms: sheet("房间数据库"),
      rules: sheet("规则库"),
      templates: sheet("回复模板库")
    });
    $("#status").textContent = `数据库已更新：${new Date().toLocaleString()}`;
    showRoomOverview();
  } catch (error) {
    if (window.KW_EMBEDDED_DB) {
      applyDatabase(window.KW_EMBEDDED_DB);
      $("#status").textContent = `云端读取失败，已使用内置备份：${error.message}`;
      showRoomOverview();
      return;
    }
    $("#status").textContent = `数据库读取失败：${error.message}`;
  }
}

function applyDatabase(database) {
  state.listings = database.listings || [];
  state.rooms = database.rooms || [];
  state.rules = (database.rules || []).filter((row) => String(row["是否启用"] || "TRUE").toUpperCase() !== "FALSE");
  state.templates = database.templates || [];
  fillListings();
}

function fetchWithTimeout(url, options, timeoutMs) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  return fetch(url, { ...options, signal: controller.signal }).finally(() => clearTimeout(timer));
}

function fillListings() {
  $("#listing").innerHTML = state.listings.map((item) => `<option value="${escapeHtml(item["房源ID"])}">${escapeHtml(item["房源名"])}</option>`).join("");
  state.listing = state.listings[0];
  fillRooms();
}

function fillRooms() {
  const rooms = currentRooms();
  $("#room").innerHTML = rooms.map((item) => `<option value="${escapeHtml(item["房间号"])}">${escapeHtml(item["房间号"])}</option>`).join("");
  state.room = rooms[0];
}

function currentRooms() {
  return state.rooms.filter((item) => !state.listing || item["房源ID"] === state.listing["房源ID"]);
}

function showRoomOverview() {
  $("#query").value = "";
  const result = roomOverview();
  applyResult(result);
  showCandidates([]);
}

function search(raw) {
  const query = String(raw || "").trim();
  if (!query) {
    showRoomOverview();
    return;
  }
  const candidates = buildCandidates(query);
  showCandidates(candidates);
  applyResult(candidates[0] || fallback(query));
}

function buildCandidates(query) {
  const results = [];
  const nq = normalize(query);
  for (const item of roomFieldCandidates()) {
    const text = normalize(`${item.key}${item.value}`);
    if (text.includes(nq) || nq.includes(text)) results.push({ ...item, score: 300 + Math.min(text.length, 80) });
  }
  for (const rule of state.rules) {
    if (!ruleApplies(rule)) continue;
    const score = scoreRule(query, rule);
    if (score <= 0) continue;
    const template = findTemplate(rule["分类"]);
    results.push({
      source: `数据库规则：${rule["分类"]}`,
      note: rule["内部备注"] || (template && template["内部备注"]) || "",
      replies: template ? fillTemplate(template) : fallbackReplies(rule["原始规则逻辑"] || rule["分类"]),
      score
    });
  }
  return results.sort((a, b) => b.score - a.score).slice(0, 12);
}

function roomFieldCandidates() {
  if (!state.room) return [];
  return Object.entries(state.room)
    .filter(([key, value]) => value && !["房源ID", "房间号", "平台显示名称/房型", "__col_18"].includes(key))
    .map(([key, value]) => ({
      source: `房间数据库：${headerName(key, "zh")}`,
      note: "已从当前房间数据库字段读取。",
      key,
      value,
      replies: roomFieldReply(key, value)
    }));
}

function roomOverview() {
  if (!state.room) return fallback("未选择房间");
  return {
    source: `房间数据库：${roomLabel("zh")}`,
    note: "页面打开默认显示当前房间的房间数据库完整信息。",
    replies: {
      zh: roomOverviewText("zh"),
      ja: roomOverviewText("ja"),
      en: roomOverviewText("en")
    }
  };
}

function roomOverviewText(language) {
  const lines = [roomLine(language)];
  for (const [key, value] of Object.entries(state.room || {})) {
    if (!value || key.startsWith("__") || key === "房源ID") continue;
    lines.push(`${headerName(key, language)}：${translateRoom(value, language)}`);
  }
  return lines.join("\n");
}

function roomFieldReply(key, value) {
  return {
    zh: `${roomLine("zh")}\n${headerName(key, "zh")}：\n${translateRoom(value, "zh")}\n如有其他問題，歡迎隨時與我們聯繫。`,
    ja: `${roomLine("ja")}\n${headerName(key, "ja")}：\n${translateRoom(value, "ja")}\nご不明な点がございましたら、お気軽にご連絡くださいませ。`,
    en: `${roomLine("en")}\n${headerName(key, "en")}:\n${translateRoom(value, "en")}\nIf you have any further questions, please feel free to contact us.`
  };
}

function fillTemplate(template) {
  const fields = {
    wifi_info: translateRoom(state.room && state.room["WiFi ID和密码"], "zh"),
    floor: translateRoom(state.room && state.room["楼层/单元"], "zh"),
    bedding: translateRoom(state.room && state.room["床具类型与数量"], "zh"),
    room_name: roomLabel("zh")
  };
  return {
    zh: applyFields(template["繁体中文回复"], fields),
    ja: applyFields(template["日语回复"], { ...fields, wifi_info: translateRoom(state.room && state.room["WiFi ID和密码"], "ja"), bedding: translateRoom(state.room && state.room["床具类型与数量"], "ja"), room_name: roomLabel("ja") }),
    en: applyFields(template["英语回复"], { ...fields, wifi_info: translateRoom(state.room && state.room["WiFi ID和密码"], "en"), bedding: translateRoom(state.room && state.room["床具类型与数量"], "en"), room_name: roomLabel("en") })
  };
}

function scoreRule(query, rule) {
  const q = normalize(query);
  const words = String(rule["关键词"] || "").split(/[,，、\n]/).map(normalize).filter(Boolean);
  let score = 0;
  for (const word of words) {
    if (q === word) score = Math.max(score, 220);
    else if (q.includes(word) || word.includes(q)) score = Math.max(score, 120);
  }
  return score ? score + Number(rule["优先级"] || 0) : 0;
}

function ruleApplies(rule) {
  const listingId = normalize(rule["房源ID"]);
  const roomNo = normalize(rule["房间号"]);
  if (listingId && !state.listing) return false;
  if (listingId && listingId !== normalize(state.listing["房源ID"])) return false;
  if (roomNo && !state.room) return false;
  if (roomNo && roomNo !== normalize(state.room["房间号"])) return false;
  if (isExclusiveRule(rule) && !state.listing) return false;
  if (state.listing && mentionsOtherListing(rule)) return false;
  return true;
}

function isExclusiveRule(rule) {
  return Boolean(normalize(rule["房源ID"]) || normalize(rule["房间号"]) || /房源专属|房间专属|专属/.test(String(rule["适用范围"] || "")) || mentionsAnyListing(rule));
}

function mentionsAnyListing(rule) {
  const text = normalize([rule["适用范围"], rule["分类"], rule["关键词"], rule["内部备注"], rule["原始规则逻辑"]].join(" "));
  return (state.listings || []).some((listing) => [listing["房源ID"], listing["房源名"], listing["平台名称关键词"]]
    .join(",")
    .split(/[,，、\n]/)
    .map(normalize)
    .filter(Boolean)
    .some((name) => text.includes(name)));
}

function mentionsOtherListing(rule) {
  const text = normalize([rule["适用范围"], rule["分类"], rule["关键词"], rule["内部备注"], rule["原始规则逻辑"]].join(" "));
  const currentId = normalize(state.listing && state.listing["房源ID"]);
  for (const listing of state.listings || []) {
    const id = normalize(listing["房源ID"]);
    if (!id || id === currentId) continue;
    const names = [listing["房源ID"], listing["房源名"], listing["平台名称关键词"]]
      .join(",")
      .split(/[,，、\n]/)
      .map(normalize)
      .filter(Boolean);
    if (names.some((name) => name && text.includes(name))) return true;
  }
  return false;
}

function findTemplate(category) {
  const key = normalize(category);
  return state.templates.find((item) => normalize(`${item["适用范围"]}${item["分类"]}`).includes(key));
}

function fallback(query) {
  return { source: "未命中数据库，已生成管家建议", note: "", replies: fallbackReplies(query), score: 1 };
}

function fallbackReplies(query) {
  return {
    zh: `您好，非常抱歉造成您的不便。\n關於「${query}」，我們會馬上確認情況並盡快回覆您。\n如有其他問題，歡迎隨時與我們聯繫。`,
    ja: `この度はご不便をおかけしてしまい、誠に申し訳ございません。\n「${query}」につきまして、すぐに状況を確認し、できるだけ早くご案内いたします。\nご不明な点がございましたら、いつでもご連絡くださいませ。`,
    en: `We sincerely apologize for the inconvenience.\nRegarding "${query}", we will check the situation right away and get back to you as soon as possible.\nIf you have any further questions, please feel free to contact us.`
  };
}

function applyResult(result) {
  state.result = result;
  $("#source").textContent = result.source;
  $("#note").textContent = result.note || "";
  $("#zh").value = result.replies.zh || "";
  $("#ja").value = result.replies.ja || "";
  $("#en").value = result.replies.en || "";
}

function showCandidates(items) {
  const box = $("#candidates");
  box.classList.toggle("show", items.length > 0);
  box.innerHTML = items.map((item, index) => `<button class="candidate" data-index="${index}">${escapeHtml(item.source)}${item.note ? `<span>${escapeHtml(item.note)}</span>` : ""}</button>`).join("");
  box.querySelectorAll(".candidate").forEach((button) => button.addEventListener("click", () => applyResult(items[Number(button.dataset.index)])));
}

function roomLine(language) {
  return language === "en" ? `Room: ${state.listing["房源名"]} Room ${state.room["房间号"]}` : language === "ja" ? `お部屋：${state.listing["房源名"]} ${state.room["房间号"]}号室` : `房源/房間：${state.listing["房源名"]} ${state.room["房间号"]}房`;
}

function roomLabel(language) {
  return language === "en" ? `${state.listing["房源名"]} Room ${state.room["房间号"]}` : language === "ja" ? `${state.listing["房源名"]} ${state.room["房间号"]}号室` : `${state.listing["房源名"]} ${state.room["房间号"]}房`;
}

function headerName(key, language) {
  const map = {
    "房间号": { zh: "房間號", ja: "部屋番号", en: "Room number" },
    "平台显示名称/房型": { zh: "平台顯示名稱/房型", ja: "掲載名/部屋タイプ", en: "Platform display name / room type" },
    "面积㎡": { zh: "面積", ja: "広さ", en: "Area" },
    "楼层/单元": { zh: "樓層/單元", ja: "階/ユニット", en: "Floor/unit" },
    "最大入住人数": { zh: "最多入住人數", ja: "最大宿泊人数", en: "Maximum occupancy" },
    "床具类型与数量": { zh: "床具類型與數量", ja: "寝具の種類と数量", en: "Bedding type and quantity" },
    "WiFi ID和密码": { zh: "WiFi ID和密碼", ja: "WiFi IDとパスワード", en: "WiFi ID and password" },
    "洗衣机": { zh: "洗衣機", ja: "洗濯機", en: "Washing machine" },
    "烘干机": { zh: "烘乾機", ja: "乾燥機", en: "Dryer" },
    "厨房": { zh: "廚房", ja: "キッチン", en: "Kitchen" },
    "电视/视频": { zh: "電視/影音", ja: "テレビ/動画サービス", en: "TV/video services" },
    "是否可开窗": { zh: "窗戶是否可開", ja: "窓の開閉", en: "Window opening" },
    "是否有电梯": { zh: "是否有電梯", ja: "エレベーター", en: "Elevator" },
    "浴室": { zh: "浴室", ja: "浴室", en: "Bathroom" },
    "厕所": { zh: "廁所", ja: "トイレ", en: "Toilet" },
    "毛巾/浴巾": { zh: "毛巾/浴巾", ja: "タオル/バスタオル", en: "Towels/bath towels" },
    "动态字段_房间设施": { zh: "房間設施", ja: "お部屋設備", en: "Room facilities" }
  };
  return (map[key] && map[key][language]) || translateRoom(key, language);
}

function translateRoom(value, language) {
  const text = String(value || "").normalize("NFKC").replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/，/g, "、").replace(/；/g, ";").replace(/：/g, ":").replace(/（/g, "(").replace(/）/g, ")");
  if (language === "zh") return applyRoomReplacements(text, [[/ダブルベッド/g, "雙人床"], [/シングルベッド/g, "單人床"], [/掛布団/g, "棉被"], [/双人床/g, "雙人床"], [/单人床/g, "單人床"], [/被子/g, "棉被"], [/枕(?!頭)/g, "枕頭"], [/房内/g, "房內"], [/烘干/g, "烘乾"], [/酱油/g, "醬油"], [/窗户/g, "窗戶"], [/换气/g, "換氣"], [/无电梯/g, "無電梯"], [/护发素/g, "護髮素"], [/洗发水/g, "洗髮水"], [/楼/g, "樓"]]);
  if (language === "ja") return toFullWidthDigits(applyRoomReplacements(text, roomReplacementsJa()).replace(/お部屋内独立洗濯機/g, "お部屋内に専用洗濯機がございます").replace(/お部屋内无乾燥機/g, "お部屋内に乾燥機はございません").replace(/是否可开窗/g, "窓の開閉").replace(/是否有エレベーター/g, "エレベーター").replace(/宿泊名数分/g, "宿泊人数分").replace(/独立/g, "専用").replace(/専用した浴室/g, "独立した浴室").replace(/専用したトイレ/g, "独立したトイレ").replace(/无/g, "なし").replace(/客人/g, "お客様").replace(/自有账号/g, "ご自身のアカウント"));
  if (language === "en") return applyRoomReplacements(text, roomReplacementsEn()).replace(/in the room独立washing machine/g, "Private washing machine in the room").replace(/in the room无dryer/g, "No dryer in the room").replace(/是否可开窗/g, "Window opening").replace(/是否有elevator/g, "Elevator").replace(/独立/g, "private").replace(/无/g, "no").replace(/客人/g, "guest").replace(/自有账号/g, "own account").replace(/\b1 double beds\b/g, "1 double bed").replace(/\b1 single beds\b/g, "1 single bed").replace(/\b1 duvets\b/g, "1 duvet").replace(/\b1 pillows\b/g, "1 pillow");
  return text;
}

function roomReplacementsJa() {
  return [[/(?:双人床|雙人床|ダブルベッド)\s*(\d+)\s*台?/g, (_m, n) => `ダブルベッド${n}台`], [/(?:单人床|單人床|シングルベッド)\s*(\d+)\s*台?/g, (_m, n) => `シングルベッド${n}台`], [/(?:被子|棉被|掛布団)\s*(\d+)\s*(?:個|个|つ)?/g, (_m, n) => `掛布団${n}つ`], [/(?:枕头|枕頭|枕)\s*(\d+)\s*(?:個|个|つ)?/g, (_m, n) => `枕${n}つ`], [/房间/g, "お部屋"], [/房内/g, "お部屋内"], [/双人床/g, "ダブルベッド"], [/單人床|单人床/g, "シングルベッド"], [/被子|棉被/g, "掛布団"], [/枕头|枕頭/g, "枕"], [/有准备晾衣架/g, "物干しラックをご用意しております"], [/附近步行约1分钟有烘干房\(投币式\)/g, "徒歩約1分の場所にコイン式乾燥室がございます"], [/有厨房/g, "キッチンがございます"], [/盐\/油\/酱油\/胡椒为独立包装/g, "塩・油・醤油・胡椒は個包装でご用意しております"], [/可看 Amazon \/ Netflix \/ Hulu \/ YouTube/g, "Amazon / Netflix / Hulu / YouTubeをご視聴いただけます"], [/需客人自有账号/g, "お客様ご自身のアカウントが必要です"], [/不可看 BS\/CS\/地上波/g, "BS/CS/地上波放送はご視聴いただけません"], [/全部窗户可打开换气/g, "すべての窓を開けて換気できます"], [/无电梯/g, "エレベーターはございません"], [/独立浴室/g, "独立した浴室がございます"], [/独立厕所/g, "独立したトイレがございます"], [/准备人数份的毛巾和浴巾/g, "宿泊人数分のタオルとバスタオルをご用意しております"], [/浴室有洗发水/g, "浴室にシャンプーがございます"], [/洗衣机/g, "洗濯機"], [/烘干机/g, "乾燥機"], [/厨房/g, "キッチン"], [/电视/g, "テレビ"], [/视频/g, "動画サービス"], [/窗户/g, "窓"], [/电梯/g, "エレベーター"], [/厕所/g, "トイレ"], [/毛巾/g, "タオル"], [/浴巾/g, "バスタオル"], [/护发素/g, "コンディショナー"], [/沐浴露/g, "ボディソープ"], [/楼/g, "階"], [/人/g, "名"]];
}

function roomReplacementsEn() {
  return [[/(?:双人床|雙人床|ダブルベッド)\s*(\d+)\s*台?/g, (_m, n) => `${n} double ${Number(n) === 1 ? "bed" : "beds"}`], [/(?:单人床|單人床|シングルベッド)\s*(\d+)\s*台?/g, (_m, n) => `${n} single ${Number(n) === 1 ? "bed" : "beds"}`], [/(?:被子|棉被|掛布団)\s*(\d+)\s*(?:個|个|つ)?/g, (_m, n) => `${n} ${Number(n) === 1 ? "duvet" : "duvets"}`], [/(?:枕头|枕頭|枕)\s*(\d+)\s*(?:個|个|つ)?/g, (_m, n) => `${n} ${Number(n) === 1 ? "pillow" : "pillows"}`], [/房间/g, "room"], [/房内/g, "in the room"], [/双人床|雙人床|ダブルベッド/g, "double bed"], [/单人床|單人床|シングルベッド/g, "single bed"], [/被子|棉被|掛布団/g, "duvet"], [/枕头|枕頭|枕/g, "pillow"], [/有准备晾衣架/g, "a drying rack is provided"], [/附近步行约1分钟有烘干房\(投币式\)/g, "there is a coin-operated dryer room about 1 minute away on foot"], [/有厨房/g, "Kitchen available"], [/盐\/油\/酱油\/胡椒为独立包装/g, "salt/oil/soy sauce/pepper are provided in individual packages"], [/可看 Amazon \/ Netflix \/ Hulu \/ YouTube/g, "Amazon / Netflix / Hulu / YouTube are available"], [/需客人自有账号/g, "guest's own account is required"], [/不可看 BS\/CS\/地上波/g, "BS/CS/terrestrial TV channels are not available"], [/全部窗户可打开换气/g, "all windows can be opened for ventilation"], [/无电梯/g, "no elevator"], [/独立浴室/g, "private bathroom"], [/独立厕所/g, "private toilet"], [/准备人数份的毛巾和浴巾/g, "towels and bath towels are prepared for the number of guests"], [/浴室有洗发水/g, "shampoo is available in the bathroom"], [/洗衣机/g, "washing machine"], [/烘干机/g, "dryer"], [/厨房/g, "kitchen"], [/电视/g, "TV"], [/视频/g, "video services"], [/窗户/g, "windows"], [/电梯/g, "elevator"], [/浴室/g, "bathroom"], [/厕所/g, "toilet"], [/毛巾/g, "towel"], [/浴巾/g, "bath towel"], [/护发素/g, "conditioner"], [/沐浴露/g, "body soap"], [/楼/g, "F"], [/人/g, " guests"], [/台|个/g, ""]];
}

function applyRoomReplacements(text, replacements) {
  let result = String(text || "");
  for (const [pattern, replacement] of replacements) result = result.replace(pattern, replacement);
  return result.split("\n").map((line) => line.trim()).filter(Boolean).join("\n");
}

function applyFields(text, fields) {
  return String(text || "").replace(/\{([^}]+)\}/g, (_match, key) => fields[key] || "");
}

function toFullWidthDigits(text) {
  return String(text || "").replace(/[0-9]/g, (digit) => "０１２３４５６７８９"[Number(digit)]);
}

function debounce(fn, wait) {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}

function escapeHtml(value) {
  return String(value || "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[char]));
}

async function copy(text) {
  await navigator.clipboard.writeText(text);
  $("#status").textContent = "已复制";
}
