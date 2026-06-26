(function () {
  "use strict";

  const DB_URL = "https://keyboard-warrior-db-1443038288.cos.ap-shanghai.myqcloud.com/%E7%BB%88%E6%9E%81%E7%89%88_%E6%95%B0%E6%8D%AE%E5%BA%93%E4%BF%AE%E6%94%B920260620.before-period-break-20260620-133222.xlsx";
  const SHORTAGE_WORDS = ["没", "没有", "没了", "用完", "用完了", "缺少", "不足", "missing", "empty", "no", "ない", "ありません", "なくなった", "足りない"];
  const SUPPLY_WORDS = ["洗发水", "洗髮水", "shampoo", "シャンプー", "护发素", "護髮素", "conditioner", "沐浴露", "bodysoap", "bodywash", "毛巾", "浴巾", "towel", "纸", "toiletpaper", "垃圾袋"];
  const state = { database: null, listing: null, room: null, matched: null };

  const $ = (selector) => document.querySelector(selector);

  document.addEventListener("DOMContentLoaded", () => {
    $("#listing").addEventListener("change", onListingChange);
    $("#room").addEventListener("change", onRoomChange);
    $("#search").addEventListener("input", debounce(runSearch, 90));
    $("#copy-ja").addEventListener("click", () => copyText($("#ja").value, "日语"));
    $("#copy-zh").addEventListener("click", () => copyText($("#zh").value, "中文"));
    $("#copy-en").addEventListener("click", () => copyText($("#en").value, "英语"));
    loadDatabase();
  });

  async function loadDatabase() {
    setStatus("正在读取信息库...");
    try {
      const database = await fetchCloudDatabase();
      state.database = database;
      setStatus(`已读取腾讯云：${new Date(database.loadedAt).toLocaleString()}`);
    } catch (error) {
      if (!window.KW_EMBEDDED_DATABASE) throw error;
      state.database = window.KW_EMBEDDED_DATABASE;
      setStatus(`网络读取失败，已使用内置数据：${new Date(state.database.loadedAt).toLocaleString()}`);
    }
    setupSelectors();
    showRoomOverview();
  }

  async function fetchCloudDatabase() {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 12000);
    try {
      const response = await fetch(`${DB_URL}?_=${Date.now()}`, { cache: "no-store", signal: controller.signal });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const bytes = new Uint8Array(await response.arrayBuffer());
      return parseWorkbookBytes(bytes, Date.now());
    } finally {
      clearTimeout(timer);
    }
  }

  function setupSelectors() {
    const listingSelect = $("#listing");
    listingSelect.innerHTML = "";
    state.database.listings.forEach((listing, index) => {
      const option = document.createElement("option");
      option.value = String(index);
      option.textContent = listing["房源名"] || listing["房源ID"] || `房源${index + 1}`;
      listingSelect.appendChild(option);
    });
    listingSelect.value = "0";
    onListingChange();
  }

  function onListingChange() {
    const index = Number($("#listing").value || 0);
    state.listing = state.database.listings[index] || null;
    const rooms = currentRooms();
    const roomSelect = $("#room");
    roomSelect.innerHTML = "";
    rooms.forEach((room, roomIndex) => {
      const option = document.createElement("option");
      option.value = String(roomIndex);
      option.textContent = room["房间号"] || room["平台显示名称/房型"] || `房间${roomIndex + 1}`;
      roomSelect.appendChild(option);
    });
    roomSelect.value = "0";
    onRoomChange();
  }

  function onRoomChange() {
    state.room = currentRooms()[Number($("#room").value || 0)] || null;
    if ($("#search").value.trim()) runSearch();
    else showRoomOverview();
  }

  function currentRooms() {
    if (!state.listing) return [];
    const listingId = normalize(state.listing["房源ID"]);
    return state.database.rooms.filter((room) => normalize(room["房源ID"]) === listingId);
  }

  function runSearch() {
    const text = cleanText($("#search").value);
    if (!text) {
      showRoomOverview();
      return;
    }
    const result = matchDatabase(text, { allowCommand: true });
    render(result.replies, result.source, result.note);
    renderCandidates(text, result);
  }

  function showRoomOverview() {
    const result = roomOverview();
    render(result.replies, result.source, result.note);
    $("#candidates").hidden = true;
    $("#candidates").innerHTML = "";
  }

  function roomOverview() {
    if (!state.listing) {
      return { source: "房间数据库：未选择房源", note: "请选择房源和房间。", replies: {
        zh: "請先選擇房源和房間。",
        ja: "施設とお部屋を選択してください。",
        en: "Please select a listing and room."
      } };
    }
    if (!state.room) {
      return { source: `房间数据库：${state.listing["房源名"]}`, note: "已选择房源，但未选择房间。", replies: {
        zh: `已選擇房源：${state.listing["房源名"]}`,
        ja: `施設を選択しました：${state.listing["房源名"]}`,
        en: `Listing selected: ${state.listing["房源名"]}`
      } };
    }
    return { source: `房间数据库：${roomLabel("zh")}`, note: "当前房间的房间数据库完整信息。", replies: {
      zh: buildRoomText("zh"),
      ja: buildRoomText("ja"),
      en: buildRoomText("en")
    } };
  }

  function matchDatabase(text, options = {}) {
    const normalized = normalize(text);
    if (!normalized) return { source: "未输入有效内容", replies: fallbackReply("") };
    const command = options.allowCommand ? matchCommandTemplate(normalized) : null;
    if (command) return command;
    const dynamic = matchDynamic(normalized);
    if (dynamic) return dynamic;
    const ranked = state.database.rules.map((rule) => ({ rule, score: scoreRule(rule, normalized) }))
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score || scopeRank(b.rule) - scopeRank(a.rule) || Number(b.rule["优先级"] || 0) - Number(a.rule["优先级"] || 0));
    if (ranked.length) {
      const rule = ranked[0].rule;
      const template = findTemplate(rule);
      return { source: `数据库规则：${rule["分类"]}`, note: rule["内部备注"] || (template && template["内部备注"]) || "", replies: template ? fillTemplate(template) : fallbackReply(rule["原始规则逻辑"] || rule["分类"]) };
    }
    return { source: "未命中数据库，已生成管家建议", replies: fallbackReply(text) };
  }

  function matchCommandTemplate(normalized) {
    const commands = state.database.commands || [];
    if (!commands.length || normalized.length > 80) return null;
    const ranked = commands.map((command) => ({ command, score: scoreCommand(command, normalized) }))
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score || Number(b.command["优先级"] || 0) - Number(a.command["优先级"] || 0));
    if (!ranked.length) return null;
    const row = ranked[0].command;
    return { source: `通用指令：${row["指令名"] || "客服短指令"}`, note: row["内部备注"] || "客服短指令生成，不绑定房源。", replies: {
      zh: row["繁体中文回复"] || "",
      ja: row["日语回复"] || "",
      en: row["英语回复"] || ""
    } };
  }

  function scoreCommand(command, normalized) {
    const words = [command["指令名"], command["搜索词/触发词"]].flatMap(splitKeywords).map((word) => normalize(word)).filter(Boolean);
    let score = 0;
    for (const word of words) {
      if (normalized === word) score = Math.max(score, 1000);
      else if (normalized.includes(word) && (word.length >= 2 || isCjkKeyword(word))) score = Math.max(score, 650 + Math.min(word.length, 40));
      else if (word.includes(normalized) && (normalized.length >= 2 || isCjkKeyword(normalized))) score = Math.max(score, 360 + Math.min(normalized.length, 40));
    }
    return score ? score + Number(command["优先级"] || 0) / 10 : 0;
  }

  function matchDynamic(normalized) {
    const fixed = [
      ["wifi", ["wifi", "无线", "网络", "internet", "ネット"]],
      ["room_area", ["面积", "面積", "size", "広さ"]],
      ["room_floor", ["楼层", "几楼", "floor", "階", "何階"]],
      ["bedding", ["床", "bed", "ベッド", "寝具"]],
      ["max_guests", ["几个人", "多少人", "capacity", "何名"]]
    ];
    for (const [scope, words] of fixed) {
      if (!words.some((word) => normalized.includes(normalize(word)))) continue;
      const template = state.database.templates.find((row) => normalize(row["适用范围"]) === normalize(scope));
      if (template) return { source: `房间数据库：${template["分类"]}`, note: template["内部备注"] || "", replies: fillTemplate(template) };
    }
    return matchRoomField(normalized);
  }

  function matchRoomField(normalized) {
    if (!state.room) return null;
    for (const [key, value] of Object.entries(state.room)) {
      if (key === "房源ID" || isBlankHeader(key) || !value) continue;
      const pieces = [key, headerName(key, "zh"), headerName(key, "ja"), headerName(key, "en"), value].flatMap(splitKeywords).map((item) => normalize(expandText(item)));
      if (!pieces.some((piece) => piece && (piece.includes(normalized) || normalized.includes(piece)))) continue;
      return { source: `房间数据库：${headerName(key, "zh")}`, note: "已从当前房间数据库字段读取。", replies: {
        zh: fieldText(key, value, "zh"),
        ja: fieldText(key, value, "ja"),
        en: fieldText(key, value, "en")
      } };
    }
    return null;
  }

  function renderCandidates(text) {
    const box = $("#candidates");
    const normalized = normalize(text);
    const candidates = [];
    const command = matchCommandTemplate(normalized);
    if (command) candidates.push({ ...command, score: 9999 });
    const dynamic = matchDynamic(normalized);
    if (dynamic) candidates.push(dynamic);
    for (const rule of state.database.rules) {
      const score = scoreRule(rule, normalized) || listRelevanceScore(rule, normalized);
      if (!score) continue;
      const template = findTemplate(rule);
      candidates.push({ source: `数据库规则：${rule["分类"]}`, note: rule["内部备注"] || (template && template["内部备注"]) || "", replies: template ? fillTemplate(template) : fallbackReply(rule["原始规则逻辑"] || rule["分类"]), score });
    }
    const unique = candidates.sort((a, b) => (b.score || 0) - (a.score || 0))
      .filter((item, index, array) => array.findIndex((candidate) => candidate.source === item.source) === index)
      .slice(0, 25);
    if (!text || !unique.length) {
      box.hidden = true;
      box.innerHTML = "";
      return;
    }
    box.hidden = false;
    box.innerHTML = `<div class="candidate-title">候选回答</div>`;
    unique.forEach((item) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "candidate";
      button.innerHTML = `<strong>${escapeHtml(item.source)}</strong>${item.note ? `<span>${escapeHtml(item.note)}</span>` : ""}`;
      button.addEventListener("click", () => render(item.replies, item.source, item.note));
      box.appendChild(button);
    });
  }

  function scoreRule(rule, normalized) {
    if (!ruleApplies(rule)) return 0;
    const category = String(rule["分类"] || "");
    const shortageSupply = hasShortage(normalized) && hasSupply(normalized);
    if (shortageSupply && /基础备品/.test(category)) return 0;
    if (/接送机/.test(category) && /^airport$|^空港$/.test(normalized)) return 0;
    let score = phraseScore(normalized, normalize(expandText(rule["分类"])), 130);
    for (const keyword of splitKeywords(rule["关键词"])) score += phraseScore(normalized, normalize(expandText(keyword)), 180);
    if (/耗品不足|缺少|不足/.test(category) && shortageSupply) score += 360;
    if (/接送机/.test(category) && ["接送机", "接机", "机场接送", "airporttransfer", "空港送迎"].some((word) => normalized.includes(normalize(word)))) score += 240;
    if (/交通|车站|駅|station/i.test(category) && ["车站", "車站", "多远", "多久", "station", "walk", "distance", "駅", "何分"].some((word) => normalized.includes(normalize(word)))) score += 240;
    return score ? score + Number(rule["优先级"] || 0) / 10 + scopeRank(rule) * 30 : 0;
  }

  function listRelevanceScore(rule, normalized) {
    if (!ruleApplies(rule)) return 0;
    const pieces = [rule["分类"], rule["关键词"], rule["内部备注"], rule["原始规则逻辑"]].flatMap(splitKeywords).map((item) => normalize(expandText(item)));
    return pieces.some((piece) => piece && (piece.includes(normalized) || normalized.includes(piece))) ? 20 + scopeRank(rule) * 5 : 0;
  }

  function phraseScore(query, keyword, base) {
    if (!query || !keyword) return 0;
    if (query === keyword) return base + 120;
    if (query.includes(keyword) && (keyword.length >= 2 || isCjkKeyword(keyword))) return Math.floor(base / 2);
    if (keyword.includes(query) && (query.length >= 2 || isCjkKeyword(query))) return Math.floor(base / 4);
    return 0;
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
    return (state.database.listings || []).some((listing) => [listing["房源ID"], listing["房源名"], listing["平台名称关键词"]]
      .flatMap(splitKeywords).map(normalize).filter(Boolean).some((name) => text.includes(name)));
  }

  function mentionsOtherListing(rule) {
    const text = normalize([rule["适用范围"], rule["分类"], rule["关键词"], rule["内部备注"], rule["原始规则逻辑"]].join(" "));
    const currentId = normalize(state.listing && state.listing["房源ID"]);
    for (const listing of state.database.listings || []) {
      const id = normalize(listing["房源ID"]);
      if (!id || id === currentId) continue;
      const names = [listing["房源ID"], listing["房源名"], listing["平台名称关键词"]].flatMap(splitKeywords).map(normalize).filter(Boolean);
      if (names.some((name) => name && text.includes(name))) return true;
    }
    return false;
  }

  function scopeRank(rule) {
    if (normalize(rule["房间号"])) return 2;
    if (normalize(rule["房源ID"]) || /房源专属|专属/.test(String(rule["适用范围"] || ""))) return 1;
    return 0;
  }

  function findTemplate(rule) {
    const category = normalize(rule["分类"]);
    return state.database.templates.find((row) => normalize(row["分类"]) === category)
      || state.database.templates.find((row) => normalize(row["适用范围"]) === category)
      || state.database.templates.find((row) => category.includes(normalize(row["分类"])) || normalize(row["分类"]).includes(category));
  }

  function fillTemplate(template) {
    const fields = dynamicFields();
    return {
      zh: applyFields(template["繁体中文回复"], fields.zh),
      ja: applyFields(template["日语回复"], fields.ja),
      en: applyFields(template["英语回复"], fields.en)
    };
  }

  function dynamicFields() {
    return { zh: baseFields("zh"), ja: baseFields("ja"), en: baseFields("en") };
  }

  function baseFields(language) {
    return {
      wifi_info: withRoom(state.room && state.room["WiFi ID和密码"], language),
      area: inlineRoom(state.room && state.room["面积㎡"], language),
      floor: inlineRoom(translateRoom(state.room && state.room["楼层/单元"], language), language),
      bedding: withRoom(translateRoom(state.room && state.room["床具类型与数量"], language), language),
      max_guests: inlineRoom(translateRoom(state.room && state.room["最大入住人数"], language), language),
      room_facilities: withRoom(facilityLines(language), language)
    };
  }

  function facilityLines(language) {
    const excluded = new Set(["房源ID", "房间号", "平台显示名称/房型", "面积㎡", "楼层/单元", "最大入住人数", "床具类型与数量", "WiFi ID和密码"]);
    return Object.entries(state.room || {})
      .filter(([key, value]) => !excluded.has(key) && !isBlankHeader(key) && value)
      .map(([key, value]) => `${headerName(key, language)}：${translateRoom(value, language)}`)
      .join("\n");
  }

  function buildRoomText(language) {
    const lines = [language === "en" ? `Room: ${roomLabel(language)}` : language === "ja" ? `お部屋：${roomLabel(language)}` : `房源/房間：${roomLabel(language)}`];
    for (const [key, value] of Object.entries(state.room || {})) {
      if (key === "房源ID" || isBlankHeader(key) || !value) continue;
      lines.push(`${headerName(key, language)}：${translateRoom(value, language)}`);
    }
    return lines.join("\n");
  }

  function fieldText(key, value, language) {
    const head = language === "en" ? `Room: ${roomLabel(language)}` : language === "ja" ? `お部屋：${roomLabel(language)}` : `房源/房間：${roomLabel(language)}`;
    const tail = language === "en" ? "If you have any further questions, please feel free to contact us." : language === "ja" ? "ご不明な点がございましたら、お気軽にご連絡くださいませ。" : "如有其他問題，歡迎隨時與我們聯繫。";
    return `${head}\n${headerName(key, language)}：\n${translateRoom(value, language)}\n${tail}`;
  }

  function roomLabel(language) {
    if (!state.room || !state.listing) return "";
    const name = state.listing["房源名"] || state.room["房源ID"] || "";
    const no = state.room["房间号"] || state.room["平台显示名称/房型"] || "";
    if (language === "en") return `${name} Room ${no}`;
    if (language === "ja") return `${name} ${no}号室`;
    return `${name} ${no}房`;
  }

  function withRoom(value, language) {
    if (!value) return language === "en" ? "Please confirm the listing and room number before replying." : language === "ja" ? "返信前に、施設名とお部屋番号をご確認ください。" : "請先確認客人的房源和房間號後再回覆。";
    const head = language === "en" ? `Room: ${roomLabel(language)}` : language === "ja" ? `お部屋：${roomLabel(language)}` : `房源/房間：${roomLabel(language)}`;
    return `${head}\n${translateRoom(value, language)}`;
  }

  function inlineRoom(value, language) {
    if (!value) return withRoom("", language);
    return `${roomLabel(language)}：${value}`;
  }

  function fallbackReply(text) {
    const topic = text || "客人的问题";
    return {
      zh: `您好，非常抱歉造成您的不便。\n關於「${topic}」，我們會馬上確認情況並盡快回覆您。\n如有其他問題，歡迎隨時與我們聯繫。`,
      ja: `この度はご不便をおかけしてしまい、誠に申し訳ございません。\n「${topic}」につきまして、すぐに状況を確認し、できるだけ早くご案内いたします。\nご不明な点がございましたら、いつでもご連絡くださいませ。`,
      en: `We sincerely apologize for the inconvenience.\nRegarding "${topic}", we will check the situation right away and get back to you as soon as possible.\nIf you have any further questions, please feel free to contact us.`
    };
  }

  function render(replies, source, note = "") {
    state.matched = { replies, source, note };
    $("#source").textContent = source || "-";
    $("#note").textContent = note || "-";
    $("#note-wrap").style.display = note ? "block" : "none";
    $("#zh").value = replies.zh || "";
    $("#ja").value = replies.ja || "";
    $("#en").value = replies.en || "";
  }

  async function copyText(text, label) {
    try {
      await navigator.clipboard.writeText(text || "");
      setStatus(`${label}已复制`);
    } catch (_) {
      setStatus(`${label}复制失败，请长按文本框手动复制`);
    }
  }

  function parseWorkbookBytes(bytes, loadedAt) {
    if (!window.XLSX) throw new Error("Excel 解析库没有加载");
    const workbook = XLSX.read(bytes, { type: "array", cellDates: false });
    const sheets = {};
    for (const name of workbook.SheetNames) sheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "", blankrows: false, raw: false });
    return {
      sheets,
      listings: rowsToObjects(sheets["房源信息"]),
      rooms: rowsToObjects(sheets["房间数据库"]),
      rules: rowsToObjects(sheets["规则库"]).filter((row) => isEnabled(row["是否启用"])),
      templates: rowsToObjects(sheets["回复模板库"]),
      commands: rowsToObjects(sheets["客服短指令模板"]).filter((row) => isEnabled(row["是否启用"])),
      loadedAt
    };
  }

  function rowsToObjects(rows = []) {
    if (!rows.length) return [];
    const headers = rows[0].map((header, index) => String(header || `__col_${index}`).trim());
    return rows.slice(1).map((row) => {
      const object = {};
      headers.forEach((header, index) => object[header] = String(row[index] || "").trim());
      return object;
    }).filter((row) => Object.values(row).some(Boolean));
  }

  function headerName(key, language) {
    const map = {
      "房间号": { zh: "房間號", ja: "部屋番号", en: "Room number" },
      "平台显示名称/房型": { zh: "平台顯示名稱/房型", ja: "表示名/部屋タイプ", en: "Platform display name / room type" },
      "面积㎡": { zh: "面積", ja: "面積", en: "Area" },
      "楼层/单元": { zh: "樓層/單元", ja: "階数/ユニット", en: "Floor / unit" },
      "最大入住人数": { zh: "最大入住人數", ja: "最大宿泊人数", en: "Maximum guests" },
      "床具类型与数量": { zh: "床具類型與數量", ja: "寝具の種類と数量", en: "Bedding type and quantity" },
      "WiFi ID和密码": { zh: "WiFi ID和密碼", ja: "WiFi IDとパスワード", en: "WiFi ID and password" },
      "洗衣机": { zh: "洗衣機", ja: "洗濯機", en: "Washing machine" },
      "烘干机": { zh: "烘乾機", ja: "乾燥機", en: "Dryer" },
      "厨房": { zh: "廚房", ja: "キッチン", en: "Kitchen" },
      "电视/视频": { zh: "電視/影音服務", ja: "テレビ/動画サービス", en: "TV / video services" },
      "是否可开窗": { zh: "是否可開窗", ja: "窓の開閉", en: "Window opening" },
      "是否有电梯": { zh: "是否有電梯", ja: "エレベーター", en: "Elevator" },
      "浴室": { zh: "浴室", ja: "浴室", en: "Bathroom" },
      "厕所": { zh: "廁所", ja: "トイレ", en: "Toilet" },
      "毛巾/浴巾": { zh: "毛巾/浴巾", ja: "タオル/バスタオル", en: "Towels / bath towels" },
      "动态字段_房间设施": { zh: "房間設施", ja: "お部屋設備", en: "Room facilities" }
    };
    return (map[key] && map[key][language]) || key;
  }

  function translateRoom(value, language) {
    const text = normalizeRoomPunctuation(String(value || ""));
    if (language === "zh") return toTraditional(text);
    if (language === "ja") return toJapanese(text);
    if (language === "en") return toEnglish(text);
    return text;
  }

  function normalizeRoomPunctuation(text) {
    return text.replace(/[；;]/g, "\n").replace(/，/g, "，").replace(/\s*\n\s*/g, "\n").trim();
  }

  function toTraditional(text) {
    return text
      .replace(/洗衣机/g, "洗衣機").replace(/烘干机/g, "烘乾機").replace(/厨房/g, "廚房")
      .replace(/没有/g, "沒有").replace(/独立/g, "獨立").replace(/准备/g, "準備")
      .replace(/步行约/g, "步行約").replace(/分钟/g, "分鐘").replace(/可打开/g, "可打開")
      .replace(/双人床/g, "雙人床").replace(/单人床/g, "單人床")
      .replace(/不可看/g, "不可看").replace(/地上波/g, "地上波");
  }

  function toJapanese(text) {
    return text
      .replace(/洗衣机/g, "洗濯機").replace(/烘干机/g, "乾燥機").replace(/厨房/g, "キッチン")
      .replace(/房内/g, "室内").replace(/独立/g, "専用").replace(/没有/g, "なし")
      .replace(/有/g, "あり").replace(/准备/g, "用意あり").replace(/晾衣架/g, "物干しラック")
      .replace(/附近步行约/g, "近くに徒歩約").replace(/分钟/g, "分").replace(/投币式/g, "コイン式")
      .replace(/盐/g, "塩").replace(/油/g, "油").replace(/酱油/g, "醤油").replace(/胡椒/g, "こしょう")
      .replace(/独立包装/g, "個包装").replace(/可看/g, "視聴可").replace(/不可看/g, "視聴不可")
      .replace(/需要客人自有账号/g, "お客様ご自身のアカウントが必要").replace(/全部窗户可打开换气/g, "すべての窓は開閉でき、換気可能です")
      .replace(/双人床/g, "ダブルベッド").replace(/单人床/g, "シングルベッド").replace(/被子/g, "掛布団").replace(/枕/g, "枕");
  }

  function toEnglish(text) {
    return text
      .replace(/洗衣机/g, "Washing machine").replace(/烘干机/g, "Dryer").replace(/厨房/g, "Kitchen")
      .replace(/房内/g, "in the room").replace(/独立/g, "private").replace(/没有/g, "not available")
      .replace(/有/g, "available").replace(/准备/g, "available").replace(/晾衣架/g, "drying rack")
      .replace(/附近步行约/g, "about ").replace(/分钟/g, " minute walk nearby").replace(/投币式/g, "coin-operated")
      .replace(/盐/g, "salt").replace(/油/g, "oil").replace(/酱油/g, "soy sauce").replace(/胡椒/g, "pepper")
      .replace(/独立包装/g, "individually packaged").replace(/可看/g, "available").replace(/不可看/g, "not available")
      .replace(/需要客人自有账号/g, "guest's own account is required").replace(/全部窗户可打开换气/g, "all windows can be opened for ventilation")
      .replace(/双人床/g, "double bed").replace(/单人床/g, "single bed").replace(/被子/g, "duvet").replace(/枕/g, "pillow");
  }

  function applyFields(template, fields) {
    return String(template || "").replace(/\{([a-zA-Z0-9_]+)\}/g, (_, key) => fields[key] || "");
  }

  function cleanText(value) { return String(value || "").replace(/\s+/g, " ").trim(); }
  function setStatus(text) { $("#status").textContent = text; }
  function isEnabled(value) { return !value || /^(true|yes|y|1|启用|是)$/i.test(String(value).trim()); }
  function splitKeywords(value) { return String(value || "").split(/[,，、;；\n/]+/).map((item) => item.trim()).filter(Boolean); }
  function normalize(value) { return String(value || "").toLowerCase().normalize("NFKC").replace(/[\s\r\n\t:_\-・,，.。!！?？/\\()[\]【】「」『』'’]/g, ""); }
  function isCjkKeyword(value) { return /[\u3400-\u9fff\u3040-\u30ff]/.test(value); }
  function isBlankHeader(key) { return !key || /^__col_/.test(key); }
  function hasShortage(value) { return SHORTAGE_WORDS.some((word) => value.includes(normalize(word))); }
  function hasSupply(value) { return SUPPLY_WORDS.some((word) => value.includes(normalize(word))); }
  function expandText(value) { return String(value || "").replace(/wi-?fi/ig, "wifi 无线 网络 internet ネット"); }
  function escapeHtml(value) { return String(value || "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;" }[char])); }
  function debounce(fn, wait) { let timer; return (...args) => { clearTimeout(timer); timer = setTimeout(() => fn(...args), wait); }; }
})();
