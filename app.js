(function () {
  "use strict";

  const DB_URL = "https://keyboard-warrior-db-1443038288.cos.ap-shanghai.myqcloud.com/%E7%BB%88%E6%9E%81%E7%89%88_%E6%95%B0%E6%8D%AE%E5%BA%93%E4%BF%AE%E6%94%B920260620.before-period-break-20260620-133222.xlsx";
  const SHORTAGE_WORDS = ["没", "没有", "没了", "用完", "缺少", "不足", "missing", "empty", "no", "ない", "ありません", "足りない"];
  const SUPPLY_WORDS = ["洗发水", "洗髮水", "shampoo", "シャンプー", "护发素", "conditioner", "沐浴露", "bodysoap", "毛巾", "towel", "纸", "toiletpaper", "垃圾袋"];
  const SYNONYMS = { wifi:["wi-fi","无线","网络","internet","ネット"], 接送机:["接机","送机","机场接送","airport transfer","空港送迎"], 洗发水:["洗髮水","shampoo","シャンプー"], 垃圾:["trash","garbage","ごみ"] };
  const state = { database:null, listing:null, room:null };
  const $ = (selector) => document.querySelector(selector);

  $("#refresh").addEventListener("click", loadDatabase);
  $("#search").addEventListener("click", search);
  $("#query").addEventListener("keydown", (event) => { if (event.key === "Enter") search(); });
  $("#listing").addEventListener("change", () => { state.listing = state.database.listings.find((item) => item["房源ID"] === $("#listing").value) || null; fillRooms(); showOverview(); });
  $("#room").addEventListener("change", () => { state.room = state.database.rooms.find((item) => item.__id === $("#room").value) || null; showOverview(); });
  document.querySelectorAll("[data-copy]").forEach((button) => button.addEventListener("click", async () => {
    const id = button.getAttribute("data-copy");
    await navigator.clipboard.writeText($(`#${id}`).value);
    setStatus(`已复制${button.textContent.replace("复制", "")}`);
  }));
  loadDatabase();

  async function loadDatabase() {
    setStatus("正在读取数据库...");
    const response = await fetch(`${DB_URL}?_=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) throw new Error(`数据库读取失败：HTTP ${response.status}`);
    state.database = parseWorkbook(new Uint8Array(await response.arrayBuffer()));
    fillListings();
    showOverview();
    setStatus(`数据库已更新：${new Date().toLocaleString()}`);
  }

  function parseWorkbook(bytes) {
    const workbook = XLSX.read(bytes, { type:"array", cellDates:false });
    const sheets = {};
    for (const name of workbook.SheetNames) sheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header:1, defval:"", blankrows:false, raw:false });
    const rooms = rowsToObjects(sheets["房间数据库"]).map((room, index) => ({ ...room, __id:`${room["房源ID"]}-${room["房间号"]}-${index}` }));
    return { sheets, listings:rowsToObjects(sheets["房源信息"]), rooms, rules:rowsToObjects(sheets["规则库"]).filter((row) => isEnabled(row["是否启用"])), templates:rowsToObjects(sheets["回复模板库"]) };
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

  function fillListings() {
    $("#listing").innerHTML = state.database.listings.map((item) => `<option value="${escapeHtml(item["房源ID"])}">${escapeHtml(item["房源名"] || item["房源ID"])}</option>`).join("");
    state.listing = state.database.listings[0] || null;
    fillRooms();
  }

  function fillRooms() {
    const rooms = state.database.rooms.filter((room) => normalize(room["房源ID"]) === normalize(state.listing && state.listing["房源ID"]));
    $("#room").innerHTML = rooms.map((room) => `<option value="${escapeHtml(room.__id)}">${escapeHtml(room["房间号"] || room["平台显示名称/房型"])}</option>`).join("");
    state.room = rooms[0] || null;
  }

  function showOverview() {
    if (!state.room) return render({ zh:"請先選擇房源和房間。", ja:"施設とお部屋を選択してください。", en:"Please select a listing and room." }, "房间数据库");
    render({ zh:buildRoomText("zh"), ja:buildRoomText("ja"), en:buildRoomText("en") }, `房间数据库：${roomLabel("zh")}`, "当前显示所选房间的全部数据库字段。");
    $("#candidates").hidden = true;
  }

  function search() {
    const text = $("#query").value.trim();
    if (!text) return showOverview();
    const result = matchDatabase(text);
    render(result.replies, result.source, result.note);
    renderCandidates(text, result);
  }

  function matchDatabase(text) {
    const normalized = normalize(text);
    const dynamic = matchDynamic(normalized);
    if (dynamic) return dynamic;
    const ranked = state.database.rules.map((rule) => ({ rule, score:scoreRule(rule, normalized) })).filter((item) => item.score > 0).sort((a,b) => b.score - a.score || scopeRank(b.rule) - scopeRank(a.rule));
    if (ranked.length) {
      const rule = ranked[0].rule;
      const template = findTemplate(rule);
      return { source:`数据库规则：${rule["分类"]}`, note:rule["内部备注"] || (template && template["内部备注"]) || "", replies:template ? fillTemplate(template) : fallback(text) };
    }
    return { source:"未命中数据库，已生成管家建议", replies:fallback(text) };
  }

  function matchDynamic(normalized) {
    const fixed = [["wifi",["wifi","无线","网络","internet","ネット"]],["room_area",["面积","面積","size","広さ"]],["room_floor",["楼层","几楼","floor","階","何階"]],["bedding",["床","bed","ベッド","寝具"]],["max_guests",["几个人","多少人","capacity","何名"]]];
    for (const [scope, words] of fixed) {
      if (!words.some((word) => normalized.includes(normalize(word)))) continue;
      const template = state.database.templates.find((row) => normalize(row["适用范围"]) === normalize(scope));
      if (template) return { source:`房间数据库：${template["分类"]}`, note:template["内部备注"] || "", replies:fillTemplate(template) };
    }
    return matchRoomField(normalized);
  }

  function matchRoomField(normalized) {
    for (const [key, value] of Object.entries(state.room || {})) {
      if (key === "房源ID" || key === "__id" || isBlankHeader(key) || !value) continue;
      const pieces = [key, headerName(key, "zh"), headerName(key, "ja"), headerName(key, "en"), value].flatMap(splitKeywords).map((item) => normalize(expandText(item)));
      if (!pieces.some((piece) => piece && (piece.includes(normalized) || normalized.includes(piece)))) continue;
      return { source:`房间数据库：${headerName(key, "zh")}`, note:"已从所选房间数据库字段读取。", replies:{ zh:fieldText(key,value,"zh"), ja:fieldText(key,value,"ja"), en:fieldText(key,value,"en") } };
    }
    return null;
  }

  function renderCandidates(text, selected) {
    const normalized = normalize(text);
    const candidates = [];
    const dynamic = matchDynamic(normalized);
    if (dynamic) candidates.push(dynamic);
    for (const rule of state.database.rules) {
      const score = scoreRule(rule, normalized) || listRelevanceScore(rule, normalized);
      if (!score) continue;
      const template = findTemplate(rule);
      candidates.push({ source:`数据库规则：${rule["分类"]}`, note:rule["内部备注"] || (template && template["内部备注"]) || "", replies:template ? fillTemplate(template) : fallback(text), score });
    }
    const unique = candidates.sort((a,b) => (b.score || 0) - (a.score || 0)).filter((item, index, array) => array.findIndex((candidate) => candidate.source === item.source) === index).slice(0, 20);
    const box = $("#candidates");
    if (!unique.length) { box.hidden = true; return; }
    box.hidden = false;
    box.innerHTML = unique.map((item, index) => `<button class="candidate" data-index="${index}" type="button"><strong>${escapeHtml(item.source)}</strong>${item.note ? `<span>${escapeHtml(item.note)}</span>` : ""}</button>`).join("");
    box.querySelectorAll("button").forEach((button) => button.addEventListener("click", () => { const item = unique[Number(button.dataset.index)]; render(item.replies, item.source, item.note); }));
  }

  function scoreRule(rule, normalized) {
    if (!ruleApplies(rule)) return 0;
    const category = String(rule["分类"] || "");
    const shortageSupply = hasShortage(normalized) && hasSupply(normalized);
    if (shortageSupply && /基础备品/.test(category)) return 0;
    let score = phraseScore(normalized, normalize(expandText(rule["分类"])), 130);
    for (const keyword of splitKeywords(rule["关键词"])) score += phraseScore(normalized, normalize(expandText(keyword)), 180);
    if (/耗品不足|缺少|不足/.test(category) && shortageSupply) score += 360;
    return score ? score + Number(rule["优先级"] || 0) / 10 + scopeRank(rule) * 30 : 0;
  }
  function listRelevanceScore(rule, normalized) { if (!ruleApplies(rule)) return 0; const pieces = [rule["分类"], rule["关键词"], rule["内部备注"], rule["原始规则逻辑"]].flatMap(splitKeywords).map((item) => normalize(expandText(item))); return pieces.some((piece) => piece && (piece.includes(normalized) || normalized.includes(piece))) ? 20 + scopeRank(rule) * 5 : 0; }
  function phraseScore(query, keyword, base) { if (!query || !keyword) return 0; if (query === keyword) return base + 120; if (query.includes(keyword) && keyword.length >= 2) return Math.floor(base / 2); if (keyword.includes(query) && query.length >= 2) return Math.floor(base / 4); return 0; }
  function ruleApplies(rule) { const listingId = normalize(rule["房源ID"]); const roomNo = normalize(rule["房间号"]); if (listingId && state.listing && listingId !== normalize(state.listing["房源ID"])) return false; if (roomNo && state.room && roomNo !== normalize(state.room["房间号"])) return false; return true; }
  function scopeRank(rule) { if (normalize(rule["房间号"])) return 2; if (normalize(rule["房源ID"]) || /房源专属|专属/.test(String(rule["适用范围"] || ""))) return 1; return 0; }
  function findTemplate(rule) { const category = normalize(rule["分类"]); return state.database.templates.find((row) => normalize(row["分类"]) === category) || state.database.templates.find((row) => normalize(row["适用范围"]) === category) || state.database.templates.find((row) => category.includes(normalize(row["分类"])) || normalize(row["分类"]).includes(category)); }
  function fillTemplate(template) { const fields = { zh:baseFields("zh"), ja:baseFields("ja"), en:baseFields("en") }; return { zh:applyFields(template["繁体中文回复"], fields.zh), ja:applyFields(template["日语回复"], fields.ja), en:applyFields(template["英语回复"], fields.en) }; }
  function baseFields(language) { return { wifi_info:withRoom(state.room && state.room["WiFi ID和密码"], language), area:inlineRoom(state.room && state.room["面积㎡"], language), floor:inlineRoom(translateRoom(state.room && state.room["楼层/单元"], language), language), bedding:withRoom(translateRoom(state.room && state.room["床具类型与数量"], language), language), max_guests:inlineRoom(translateRoom(state.room && state.room["最大入住人数"], language), language), room_facilities:withRoom(facilityLines(language), language) }; }
  function facilityLines(language) { const excluded = new Set(["房源ID","房间号","平台显示名称/房型","面积㎡","楼层/单元","最大入住人数","床具类型与数量","WiFi ID和密码","__id"]); return Object.entries(state.room || {}).filter(([key,value]) => !excluded.has(key) && !isBlankHeader(key) && value).map(([key,value]) => `${headerName(key, language)}：${translateRoom(value, language)}`).join("\n"); }
  function buildRoomText(language) { const lines = [language === "en" ? `Room: ${roomLabel(language)}` : language === "ja" ? `お部屋：${roomLabel(language)}` : `房源/房間：${roomLabel(language)}`]; for (const [key,value] of Object.entries(state.room || {})) { if (key === "房源ID" || key === "__id" || isBlankHeader(key) || !value) continue; lines.push(`${headerName(key, language)}：${translateRoom(value, language)}`); } return lines.join("\n"); }
  function fieldText(key, value, language) { const head = language === "en" ? `Room: ${roomLabel(language)}` : language === "ja" ? `お部屋：${roomLabel(language)}` : `房源/房間：${roomLabel(language)}`; const tail = language === "en" ? "If you have any further questions, please feel free to contact us." : language === "ja" ? "ご不明な点がございましたら、お気軽にご連絡くださいませ。" : "如有其他問題，歡迎隨時與我們聯繫。"; return `${head}\n${headerName(key, language)}：\n${translateRoom(value, language)}\n${tail}`; }
  function roomLabel(language) { const name = state.listing && (state.listing["房源名"] || state.listing["房源ID"]); const no = state.room && (state.room["房间号"] || state.room["平台显示名称/房型"]); if (language === "en") return `${name} Room ${no}`; if (language === "ja") return `${name} ${no}号室`; return `${name} ${no}房`; }
  function withRoom(value, language) { if (!value) return language === "en" ? "Please select the listing and room first." : language === "ja" ? "先に施設とお部屋を選択してください。" : "請先選擇房源和房間。"; const head = language === "en" ? `Room: ${roomLabel(language)}` : language === "ja" ? `お部屋：${roomLabel(language)}` : `房源/房間：${roomLabel(language)}`; return `${head}\n${translateRoom(value, language)}`; }
  function inlineRoom(value, language) { if (!value) return withRoom("", language); return `${roomLabel(language)}：${value}`; }
  function render(replies, source, note = "") { $("#source").textContent = source || "-"; $("#note").textContent = note || "-"; $("#note-wrap").style.display = note ? "grid" : "none"; $("#zh").value = replies.zh || ""; $("#ja").value = replies.ja || ""; $("#en").value = replies.en || ""; }
  function fallback(text) { return { zh:`您好，非常抱歉造成您的不便。\n關於「${text}」，我們會馬上確認情況並盡快回覆您。\n如有其他問題，歡迎隨時與我們聯繫。`, ja:"この度はご不便をおかけしてしまい、誠に申し訳ございません。\nお問い合わせいただいた内容につきまして、すぐに状況を確認し、できるだけ早くご案内いたします。\nご不明な点がございましたら、いつでもご連絡くださいませ。", en:"We sincerely apologize for the inconvenience.\nRegarding your inquiry, we will check the situation right away and get back to you as soon as possible.\nIf you have any further questions, please feel free to contact us." }; }
  function headerName(key, language) { const map = {"房间号":{zh:"房間號",ja:"部屋番号",en:"Room number"},"平台显示名称/房型":{zh:"平台顯示名稱/房型",ja:"掲載名/部屋タイプ",en:"Platform display name / room type"},"面积㎡":{zh:"面積",ja:"広さ",en:"Area"},"楼层/单元":{zh:"樓層/單元",ja:"階/ユニット",en:"Floor/unit"},"最大入住人数":{zh:"最多入住人數",ja:"最大宿泊人数",en:"Maximum occupancy"},"床具类型与数量":{zh:"床具類型與數量",ja:"寝具の種類と数量",en:"Bedding type and quantity"},"WiFi ID和密码":{zh:"WiFi ID和密碼",ja:"WiFi IDとパスワード",en:"WiFi ID and password"},"洗衣机":{zh:"洗衣機",ja:"洗濯機",en:"Washing machine"},"烘干机":{zh:"烘乾機",ja:"乾燥機",en:"Dryer"},"厨房":{zh:"廚房",ja:"キッチン",en:"Kitchen"},"电视/视频":{zh:"電視/影音",ja:"テレビ/動画サービス",en:"TV/video services"},"是否可开窗":{zh:"窗戶是否可開",ja:"窓の開閉",en:"Window opening"},"是否有电梯":{zh:"是否有電梯",ja:"エレベーター",en:"Elevator"},"浴室":{zh:"浴室",ja:"浴室",en:"Bathroom"},"厕所":{zh:"廁所",ja:"トイレ",en:"Toilet"},"毛巾/浴巾":{zh:"毛巾/浴巾",ja:"タオル/バスタオル",en:"Towels/bath towels"},"动态字段_房间设施":{zh:"房間設施",ja:"お部屋設備",en:"Room facilities"}}; return (map[key] && map[key][language]) || key; }
  function translateRoom(value, language) { let text = String(value || ""); const common = [[/１/g,"1"],[/２/g,"2"],[/３/g,"3"],[/４/g,"4"]]; const sets = { zh:[[/ダブルベッド/g,"雙人床"],[/シングルベッド/g,"單人床"],[/掛布団/g,"棉被"],[/枕/g,"枕頭"],[/つ/g,"個"],[/台/g,"張"],[/房内/g,"房內"],[/烘干/g,"烘乾"],[/楼/g,"樓"]], ja:[[/房内独立洗衣机/g,"お部屋内に専用洗濯機がございます"],[/房内无烘干机/g,"お部屋内に乾燥機はございません"],[/有厨房/g,"キッチンがございます"],[/无电梯/g,"エレベーターはございません"],[/独立浴室/g,"独立した浴室がございます"],[/独立厕所/g,"独立したトイレがございます"],[/楼/g,"階"]], en:[[/ダブルベッド/g,"double bed"],[/シングルベッド/g,"single bed"],[/掛布団/g,"duvet"],[/枕/g,"pillow"],[/つ/g,""],[/台/g,""],[/房内独立洗衣机/g,"Private washing machine in the room"],[/房内无烘干机/g,"No dryer in the room"],[/有厨房/g,"Kitchen available"],[/无电梯/g,"No elevator"],[/独立浴室/g,"Private bathroom"],[/独立厕所/g,"Private toilet"],[/楼/g,"F"],[/人/g," guests"]] }; for (const [pattern,replacement] of [...common, ...(sets[language] || [])]) text = text.replace(pattern,replacement); if (language === "en") text = text.replace(/double bed(\d+)/g, "$1 double bed").replace(/pillow(\d+)/g, "$1 pillow"); return text; }
  function applyFields(text, fields) { return String(text || "").replace(/\{([^}]+)\}/g, (_match, key) => fields[key] || `请先确认${key}`); }
  function expandText(text) { let expanded = ` ${text || ""} `; const normalized = normalize(text); Object.entries(SYNONYMS).forEach(([main, words]) => { if ([main, ...words].some((word) => normalized.includes(normalize(word)))) expanded += ` ${main} ${words.join(" ")} `; }); return expanded; }
  function hasShortage(value) { return SHORTAGE_WORDS.some((word) => value.includes(normalize(word))); }
  function hasSupply(value) { return SUPPLY_WORDS.some((word) => value.includes(normalize(word))); }
  function isEnabled(value) { return !value || /^(true|yes|y|1|启用|是)$/i.test(String(value).trim()); }
  function splitKeywords(value) { return String(value || "").split(/[,，、;；\n/]+/).map((item) => item.trim()).filter(Boolean); }
  function normalize(value) { return String(value || "").toLowerCase().normalize("NFKC").replace(/[\s\r\n\t:_\-・,，.。!！?？/\\()[\]【】「」『』'’]/g, ""); }
  function isBlankHeader(key) { return /^__col_\d+(?:_\d+)?$/.test(String(key || "")); }
  function escapeHtml(value) { return String(value || "").replace(/[&<>"']/g, (char) => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", "\"":"&quot;", "'":"&#39;" })[char]); }
  function setStatus(text) { $("#status").textContent = text; }
})();
