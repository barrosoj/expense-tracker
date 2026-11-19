(function(){
  "use strict";

  // Local data key. Without login: the legacy key. With login: one key per user
  // ("<legacy>:<userId>") so two users on the same device never see each other's data.
  var LEGACY_STORAGE_KEY = "quadroTarefasOffline_v1";
  var LEGACY_CLAIMED_KEY = "quadroTarefasOffline_v1_claimedBy";
  var STORAGE_KEY = LEGACY_STORAGE_KEY;
  var currentUserId = null;
  var IDB_NAME = "quadroTarefasOfflineDocs";
  var IDB_STORE = "documents";

  var STATUSES = [
    { id:"todo",  label:"A Fazer" },
    { id:"doing", label:"Em Curso" },
    { id:"done",  label:"Concluído" }
  ];
  var PALETTE = ["#2F6659","#8A5A2E","#3D5A80","#6B4E9B","#96682B","#4C6B3E","#8A3E52","#3E6B86"];
  var PRIORITY_ORDER = { alta:0, media:1, baixa:2 };
  var PRIORITY_LABEL = { alta:"Alta", media:"Média", baixa:"Baixa" };
  var MONTHS = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
  var FULL_MONTHS = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  var PROJECT_STATUS_LABEL = { ativo:"Ativo", pausado:"Pausado", concluido:"Concluído" };
  var PROJECT_STATUS_ORDER = { ativo:0, pausado:1, concluido:2 };
  // Same categories (and order) as the columns of the Despesas.xlsx monthly map.
  var EXPENSE_CATEGORIES = ["viatura","material","combustivel","almocos","jantares","refeicoes","estadias","parques","portagens","outros"];
  var EXPENSE_CATEGORY_LABEL = {
    viatura:"Viatura", material:"Material escritório", combustivel:"Combustível", almocos:"Almoços", jantares:"Jantares",
    refeicoes:"Refeições gerais", estadias:"Estadias", parques:"Parques", portagens:"Portagens", outros:"Outros"
  };
  // Categories used before the Excel columns: shown and exported as their new equivalent.
  var LEGACY_EXPENSE_CATEGORY = { deslocacoes:"outros", alojamento:"estadias" };
  function expenseCategory(x){
    var c = x && x.category;
    c = LEGACY_EXPENSE_CATEGORY[c] || c;
    return EXPENSE_CATEGORIES.indexOf(c) >= 0 ? c : "outros";
  }

  var DEFAULT_BRANDS = ["Kennametal","Bass","Arno","GAIt"];

  var currentView = "home"; // "home" | "tasks" | "meetings" | "projects" | "expenses"

  // Sales/purchase values imported from a file: { "AAAA-MM": { "Marca": { venda:n, compra:n } } }.
  // Kept on this device only (not synced), under a per-user key.
  var salesData = {};
  var mySales = {}; // my own sales: { "AAAA-MM": n|null }
  var homeMonth = null;

  var clients = {};
  var tasks = {};
  var meetings = {};
  var projects = {};
  var expenses = {};
  var activeClientFilter = null;

  var editingTaskId = null;
  var editingMeetingId = null;
  var editingProjectId = null;
  var editingExpenseId = null;
  var actionsDraft = [];
  var documentsDraft = [];
  var documentsPendingCleanup = [];
  var dragTaskId = null;
  var currentMonth = null;
  var dayModalDate = null;
  var applyingSyncedData = false;

  // ---------- utils ----------
  function el(tag, cls, html){
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (html !== undefined) e.innerHTML = html;
    return e;
  }
  function escapeHtml(s){
    return String(s == null ? "" : s).replace(/[&<>"']/g, function(c){
      return { "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" }[c];
    });
  }
  function todayStr(){
    var d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
  }
  function daysAgoStr(n){
    var d = new Date();
    d.setDate(d.getDate() - n);
    return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0");
  }
  function formatDue(due){
    if (!due) return null;
    var today = todayStr();
    if (due === today) return { text:"hoje", cls:"due-today" };
    var d = new Date(due + "T00:00:00");
    var t = new Date(today + "T00:00:00");
    var diffDays = Math.round((d - t) / 86400000);
    var label = d.getDate() + " " + MONTHS[d.getMonth()];
    if (diffDays < 0) return { text:"atrasado · " + label, cls:"due-overdue" };
    if (diffDays === 1) return { text:"amanhã", cls:"" };
    return { text:label, cls:"" };
  }
  function formatDateShort(dateStr){
    if (!dateStr) return "—";
    var d = new Date(dateStr + "T00:00:00");
    if (isNaN(d.getTime())) return dateStr;
    return d.getDate() + " " + MONTHS[d.getMonth()] + " " + d.getFullYear();
  }
  function clientColor(id){ return (id && clients[id]) ? clients[id].color : null; }
  function clientName(id){ return (id && clients[id]) ? clients[id].name : null; }
  function nextPaletteColor(){ return PALETTE[Object.keys(clients).length % PALETTE.length]; }
  function genId(){ return "id" + Date.now().toString(36) + Math.random().toString(36).slice(2,8); }
  function formatBytes(n){
    if (n == null) return "";
    if (n < 1024) return n + " B";
    if (n < 1024*1024) return (n/1024).toFixed(0) + " KB";
    return (n/(1024*1024)).toFixed(1) + " MB";
  }
  // "1234.5" → "1 234,50 €". Built by hand because pt-PT's toLocaleString leaves
  // four-digit amounts ungrouped ("8300,10 €"), which looks inconsistent next to "12 450,75 €".
  function formatCurrency(n){
    var v = Number(n) || 0;
    var parts = Math.abs(v).toFixed(2).split(".");
    var intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ");
    var negative = v < 0 && parts.join("") !== "000";
    return (negative ? "-" : "") + intPart + "," + parts[1] + " €";
  }
  function parseAmountInput(str){
    if (!str) return 0;
    var cleaned = String(str).replace(/[€\s]/g, "").replace(/\./g, "").replace(",", ".");
    var n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  }

  // ---------- IndexedDB (document blobs) ----------
  var idbPromise = null;
  function idbOpen(){
    if (idbPromise) return idbPromise;
    idbPromise = new Promise(function(resolve, reject){
      if (!window.indexedDB){ reject(new Error("IndexedDB indisponível")); return; }
      var req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = function(e){
        var idb = e.target.result;
        if (!idb.objectStoreNames.contains(IDB_STORE)) idb.createObjectStore(IDB_STORE, { keyPath:"id" });
      };
      req.onsuccess = function(e){ resolve(e.target.result); };
      req.onerror = function(e){ reject(e.target.error); };
    });
    return idbPromise;
  }
  function idbPut(doc){
    return idbOpen().then(function(idb){
      return new Promise(function(resolve, reject){
        var tx = idb.transaction(IDB_STORE, "readwrite");
        tx.objectStore(IDB_STORE).put(doc);
        tx.oncomplete = function(){ resolve(); };
        tx.onerror = function(){ reject(tx.error); };
      });
    });
  }
  function idbGet(id){
    return idbOpen().then(function(idb){
      return new Promise(function(resolve, reject){
        var tx = idb.transaction(IDB_STORE, "readonly");
        var req = tx.objectStore(IDB_STORE).get(id);
        req.onsuccess = function(){ resolve(req.result || null); };
        req.onerror = function(){ reject(req.error); };
      });
    });
  }
  function idbDelete(id){
    return idbOpen().then(function(idb){
      return new Promise(function(resolve, reject){
        var tx = idb.transaction(IDB_STORE, "readwrite");
        tx.objectStore(IDB_STORE).delete(id);
        tx.oncomplete = function(){ resolve(); };
        tx.onerror = function(){ reject(tx.error); };
      });
    });
  }

  // ---------- storage ----------
  function defaultSeed(){
    return {
      clients: {},
      tasks: {},
      meetings: {},
      projects: {},
      expenses: {}
    };
  }

  // Data saved before login existed lives under the legacy key. Offer it once to the first
  // user who signs in on this device; afterwards it's marked as claimed and never offered again.
  function takeLegacyDataIfWanted(){
    if (!currentUserId) return null;
    var legacy = null, claimedBy = null;
    try {
      legacy = localStorage.getItem(LEGACY_STORAGE_KEY);
      claimedBy = localStorage.getItem(LEGACY_CLAIMED_KEY);
    } catch(e){ return null; }
    if (!legacy || claimedBy) return null;
    var useIt = window.confirm(
      "Este navegador tem dados guardados antes de existir login.\n\n" +
      "Associar esses dados a esta conta?\n" +
      "OK = sim, ficam nesta conta.  Cancelar = não, esta conta começa vazia (os dados ficam guardados para outra conta)."
    );
    if (!useIt) return null;
    try { localStorage.setItem(LEGACY_CLAIMED_KEY, currentUserId); } catch(e){}
    return legacy;
  }

  function loadStore(){
    var raw = null;
    try { raw = localStorage.getItem(STORAGE_KEY); } catch(e){}
    if (!raw) raw = takeLegacyDataIfWanted();
    if (raw){
      try {
        var parsed = JSON.parse(raw);
        clients = parsed.clients || {};
        tasks = parsed.tasks || {};
        meetings = parsed.meetings || {};
        projects = parsed.projects || {};
        expenses = parsed.expenses || {};
        return;
      } catch(e){ /* fall through to seed */ }
    }
    var seed = defaultSeed();
    clients = seed.clients;
    tasks = seed.tasks;
    meetings = seed.meetings;
    projects = seed.projects;
    expenses = seed.expenses;
    saveStore();
  }

  function saveStore(){
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ clients:clients, tasks:tasks, meetings:meetings, projects:projects, expenses:expenses }));
      if (!applyingSyncedData && window.QuadroSync && currentUserId){
        window.QuadroSync.localChanged();
      }
      scheduleExcelSync();
    } catch(e){
      var warn = document.getElementById("storageWarning");
      if (warn){
        warn.hidden = false;
        warn.querySelector("span").textContent = "Não foi possível guardar as alterações neste navegador (armazenamento indisponível ou cheio). Usa \"Exportar cópia\" para não perderes o trabalho.";
      }
    }
  }

  // ---------- view helpers ----------
  function currentMap(){
    if (currentView === "tasks") return tasks;
    if (currentView === "meetings") return meetings;
    if (currentView === "projects") return projects;
    return expenses;
  }

  // ---------- home: sales & purchases ----------
  // Values are read from the Excel files each time; earlier versions kept a copy in this
  // browser, which is removed so no figures stay stored outside those files.
  function forgetStoredSalesCopy(){
    try { localStorage.removeItem(STORAGE_KEY + ":vendasCompras"); } catch(e){}
  }

  function brandList(){
    var seen = {};
    var list = [];
    function add(b){
      var k = b.toLowerCase();
      if (!seen[k]){ seen[k] = true; list.push(b); }
    }
    DEFAULT_BRANDS.forEach(add);
    Object.keys(salesData).sort().forEach(function(ym){
      Object.keys(salesData[ym] || {}).forEach(add);
    });
    return list;
  }
  function findBrandKey(monthData, brand){
    if (!monthData || !brand) return null;
    var target = brand.toLowerCase();
    for (var b in monthData) if (b.toLowerCase() === target) return b;
    return null;
  }

  function normalizeHeader(h){
    return String(h || "").trim().toLowerCase()
      .normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/[^a-z]/g, "");
  }
  function splitCsvLine(line, sep){
    var out = [], cur = "", inQ = false;
    for (var i = 0; i < line.length; i++){
      var c = line[i];
      if (inQ){
        if (c === '"'){
          if (line[i+1] === '"'){ cur += '"'; i++; }
          else inQ = false;
        } else cur += c;
      } else if (c === '"') inQ = true;
      else if (c === sep){ out.push(cur); cur = ""; }
      else cur += c;
    }
    out.push(cur);
    return out.map(function(s){ return s.trim(); });
  }
  // Accepts "1.234,56", "1234,56", "1234.56", "1,234.56" and "€ 1 234,56".
  function parseFileAmount(str){
    if (typeof str === "number") return isFinite(str) ? str : null;
    var s = String(str == null ? "" : str).replace(/[€\s ]/g, "");
    if (!s) return null;
    var lastComma = s.lastIndexOf(","), lastDot = s.lastIndexOf(".");
    if (lastComma > -1 && lastDot > -1){
      if (lastComma > lastDot) s = s.replace(/\./g, "").replace(",", ".");
      else s = s.replace(/,/g, "");
    } else if (lastComma > -1){
      s = s.replace(/\./g, "").replace(",", ".");
    } else if (lastDot > -1 && /^-?\d{1,3}(\.\d{3})+$/.test(s)){
      s = s.replace(/\./g, "");
    }
    var n = parseFloat(s);
    return isNaN(n) ? null : n;
  }
  // "2026-11", "2026/11", "11/2026", "11-2026", "2026-11-05", "05/11/2026" → "2026-11".
  function parseFileMonth(str){
    if (typeof str === "number" && str > 20000 && str < 80000){
      // Excel date serial number (days since 1899-12-30).
      var d = new Date(Date.UTC(1899, 11, 30) + str * 86400000);
      return d.getUTCFullYear() + "-" + String(d.getUTCMonth()+1).padStart(2,"0");
    }
    var s = String(str == null ? "" : str).trim();
    var m;
    if ((m = s.match(/^(\d{4})[-\/.](\d{1,2})(?:[-\/.]\d{1,2})?$/))) return m[1] + "-" + m[2].padStart(2,"0");
    if ((m = s.match(/^(\d{1,2})[-\/.](\d{4})$/))) return m[2] + "-" + m[1].padStart(2,"0");
    if ((m = s.match(/^\d{1,2}[-\/.](\d{1,2})[-\/.](\d{4})$/))) return m[2] + "-" + m[1].padStart(2,"0");
    return null;
  }

  function isTotalLabel(s){ return /^(sub)?total/.test(normalizeHeader(s)); }

  // "Mes01", "Mês 1", "Jan", "janeiro" → 1..12; anything else → 0.
  function headerMonthNumber(h){
    var n = normalizeHeader(h);
    var m = String(h == null ? "" : h).trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, "").match(/^mes0?(\d{1,2})$/);
    if (m){ var k = parseInt(m[1], 10); return (k >= 1 && k <= 12) ? k : 0; }
    for (var i = 0; i < 12; i++){
      var full = FULL_MONTHS[i].normalize("NFD").replace(/[̀-ͯ]/g, "");
      if (n === full || n === MONTHS[i]) return i + 1;
    }
    return 0;
  }

  // Finds a header row with at least 3 month columns ("Mes01"…"Mes12", "Jan"…).
  // Returns { headerIdx, monthCols:{ 1:colIndex, … } } or null.
  function findMonthHeader(rows){
    for (var h = 0; h < Math.min(rows.length, 20); h++){
      var monthCols = {}, nMonths = 0;
      for (var idx = 0; idx < rows[h].length; idx++){
        var k = headerMonthNumber(rows[h][idx]);
        if (k && monthCols[k] === undefined){ monthCols[k] = idx; nMonths++; }
      }
      if (nMonths >= 3) return { headerIdx:h, monthCols:monthCols };
    }
    return null;
  }
  function nonEmptyRows(table){
    var rows = table.filter(function(r){
      return r && r.some(function(c){ return c !== null && c !== undefined && String(c).trim() !== ""; });
    });
    if (rows.length < 2) throw new Error("O ficheiro não tem linhas de dados.");
    return rows;
  }
  function cellText(v){ return String(v == null ? "" : v).trim(); }

  // Purchases per brand, in one of two layouts:
  //  - one column per month (ERP export): Descricao | Mes01 … Mes12 | Total — one row per brand,
  //    values for the current year;
  //  - one row per value: Mês | Marca | Compra.
  // Returns { data:{ "AAAA-MM": { marca: n|null } }, rows, year? }.
  function parsePurchaseRows(table){
    var rows = nonEmptyRows(table);
    var found = findMonthHeader(rows);
    if (found){
      var header = rows[found.headerIdx];
      var cols = Object.keys(found.monthCols).map(function(k){ return found.monthCols[k]; });
      var brandCol = -1;
      for (var i = 0; i < header.length; i++){
        if (["marca","brand","descricao","fornecedor","artigo"].indexOf(normalizeHeader(header[i])) > -1){ brandCol = i; break; }
      }
      if (brandCol < 0) brandCol = Math.min.apply(null, cols) > 0 ? 0 : -1;
      if (brandCol < 0) throw new Error("Não encontrei a coluna das marcas.");
      var year = todayStr().slice(0,4);
      var data = {}, count = 0;
      for (var r = found.headerIdx + 1; r < rows.length; r++){
        var brand = cellText(rows[r][brandCol]);
        if (!brand || isTotalLabel(brand)) continue;
        for (var k = 1; k <= 12; k++){
          if (found.monthCols[k] === undefined) continue;
          var ym = year + "-" + String(k).padStart(2,"0");
          data[ym] = data[ym] || {};
          // An empty cell clears the value: the Excel file is the source of truth.
          data[ym][brand] = parseFileAmount(rows[r][found.monthCols[k]]);
        }
        count++;
      }
      if (!count) throw new Error("Nenhuma marca encontrada no ficheiro.");
      return { data:data, rows:count, year:year };
    }

    var headers = rows[0].map(normalizeHeader);
    function col(names){
      for (var j = 0; j < headers.length; j++) if (names.indexOf(headers[j]) > -1) return j;
      return -1;
    }
    var iMonth = col(["mes","month","periodo","data","date"]);
    var iBrand = col(["marca","brand","fornecedor","descricao"]);
    var iBuy = col(["compra","compras","valorcompra","valordecompra","valor","purchase","purchases"]);
    if (iBrand < 0) throw new Error("Não encontrei a coluna das marcas (\"Marca\" ou \"Descricao\").");
    if (iBuy < 0) throw new Error("Não encontrei os valores (colunas \"Mes01…Mes12\" ou \"Compra\").");
    var fallbackMonth = todayStr().slice(0,7);
    var out = {}, n = 0;
    for (var x = 1; x < rows.length; x++){
      var b = cellText(rows[x][iBrand]);
      var m = iMonth > -1 ? parseFileMonth(rows[x][iMonth]) : fallbackMonth;
      var v = parseFileAmount(rows[x][iBuy]);
      if (!b || isTotalLabel(b) || !m || v === null) continue;
      out[m] = out[m] || {};
      var key = findBrandKey(out[m], b) || b;
      out[m][key] = (out[m][key] || 0) + v; // repeated brand/month rows are added together
      n++;
    }
    if (!n) throw new Error("Nenhuma linha válida encontrada.");
    return { data:out, rows:n };
  }

  // My own sales: one value per month.
  //  - one column per month: uses the "Total" row if there is one, otherwise adds up all rows;
  //  - one row per value: Mês | Venda (rows of the same month are added together).
  // Returns { data:{ "AAAA-MM": n|null }, year? }.
  function parseMySalesRows(table){
    var rows = nonEmptyRows(table);
    var found = findMonthHeader(rows);
    if (found){
      var year = todayStr().slice(0,4);
      var body = rows.slice(found.headerIdx + 1);
      var labelCol = Math.min.apply(null, Object.keys(found.monthCols).map(function(k){ return found.monthCols[k]; })) > 0 ? 0 : -1;
      var totalRow = labelCol < 0 ? null : body.filter(function(r){ return isTotalLabel(cellText(r[labelCol])); })[0];
      var useRows = totalRow ? [totalRow] : body;
      var data = {};
      for (var k = 1; k <= 12; k++){
        if (found.monthCols[k] === undefined) continue;
        var sum = null;
        useRows.forEach(function(r){
          var v = parseFileAmount(r[found.monthCols[k]]);
          if (v !== null) sum = (sum || 0) + v;
        });
        data[year + "-" + String(k).padStart(2,"0")] = sum; // empty month → no value
      }
      return { data:data, year:year };
    }

    var headers = rows[0].map(normalizeHeader);
    function col(names){
      for (var j = 0; j < headers.length; j++) if (names.indexOf(headers[j]) > -1) return j;
      return -1;
    }
    var iMonth = col(["mes","month","periodo","data","date"]);
    var iSale = col(["venda","vendas","valorvenda","valordevenda","valor","total","sale","sales"]);
    if (iMonth < 0 || iSale < 0) throw new Error("Não encontrei os valores (colunas \"Mes01…Mes12\", ou \"Mês\" e \"Venda\").");
    var out = {}, n = 0;
    for (var x = 1; x < rows.length; x++){
      var m = parseFileMonth(rows[x][iMonth]);
      var v = parseFileAmount(rows[x][iSale]);
      if (!m || v === null) continue;
      out[m] = (out[m] || 0) + v;
      n++;
    }
    if (!n) throw new Error("Nenhuma linha válida encontrada.");
    return { data:out };
  }

  function parseCsvText(text){
    var lines = String(text).replace(/^﻿/, "").split(/\r\n|\n|\r/).filter(function(l){ return l.trim() !== ""; });
    if (!lines.length) return [];
    var first = lines[0];
    var sep = [";", "\t", ","].reduce(function(best, s){
      return first.split(s).length > first.split(best).length ? s : best;
    }, ";");
    return lines.map(function(l){ return splitCsvLine(l, sep); });
  }

  // ---------- values from the Compras / Vendas folders (view only) ----------
  // The Excel files live in the "Compras" and "Vendas" folders next to the app on this computer
  // (they are git-ignored, so never on GitHub). The app only reads them; it never changes or sends them.
  // They are found only when the app is opened from this computer (local server on the project folder).
  var DATA_FOLDERS = { compra:["Compras","Compra"], venda:["Vendas","Venda"] };
  var KIND_LABEL = { compra:"compras", venda:"vendas" };
  var folderLoad = null;

  function setHomeImportStatus(kind, message, isError){
    var s = document.getElementById(kind === "venda" ? "homeSaleStatus" : "homePurchaseStatus");
    s.textContent = message || "";
    s.classList.toggle("error", !!isError);
    // The status lines only show when a file could not be read.
    document.getElementById("homeFilesStatus").hidden =
      !document.querySelector("#homeFilesStatus .home-import-status.error");
  }

  function isDataFileName(name){ return /\.(xlsx|csv)$/i.test(name) && name.indexOf("~$") !== 0; }

  // Looks for the first Excel/CSV file in the kind's folder. Uses the folder listing when the
  // local server provides one, else tries "<Pasta>/<Pasta>.xlsx" (e.g. Compras/Compras.xlsx).
  function findFolderFile(kind){
    var folders = DATA_FOLDERS[kind];
    function tryFolder(i){
      if (i >= folders.length) return Promise.resolve(null);
      var folder = folders[i];
      return fetch(folder + "/", { cache:"no-store" }).then(function(res){
        if (!res.ok || (res.headers.get("content-type") || "").indexOf("text/html") < 0) return null;
        return res.text().then(function(html){
          var doc = new DOMParser().parseFromString(html, "text/html");
          var names = Array.prototype.map.call(doc.querySelectorAll("a[href]"), function(a){
            var href = a.getAttribute("href").split(/[?#]/)[0];
            try { href = decodeURIComponent(href); } catch(e){}
            return href.split("/").pop();
          }).filter(isDataFileName).sort();
          return names.length ? folder + "/" + names[0] : null;
        });
      }).catch(function(){ return null; }).then(function(found){
        if (found) return found;
        var guess = folder + "/" + folder + ".xlsx";
        return fetch(guess, { method:"HEAD", cache:"no-store" })
          .then(function(res){ return res.ok ? guess : null; })
          .catch(function(){ return null; });
      }).then(function(found){ return found || tryFolder(i + 1); });
    }
    return tryFolder(0);
  }

  function readFolderFile(path){
    return fetch(path.split("/").map(encodeURIComponent).join("/"), { cache:"no-store" }).then(function(res){
      if (!res.ok) throw new Error("não foi possível abrir " + path);
      if (/\.csv$/i.test(path)) return res.text().then(parseCsvText);
      if (!window.QuadroXlsx) throw new Error("leitor de Excel indisponível");
      return res.arrayBuffer().then(window.QuadroXlsx.read).then(function(r){ return r.rows; });
    });
  }

  function applyParsed(kind, parsed){
    Object.keys(parsed.data).forEach(function(ym){
      if (kind === "venda"){ mySales[ym] = parsed.data[ym]; return; }
      salesData[ym] = salesData[ym] || {};
      Object.keys(parsed.data[ym]).forEach(function(brand){
        var key = findBrandKey(salesData[ym], brand) || brand;
        salesData[ym][key] = { compra:parsed.data[ym][brand] };
      });
    });
  }

  function loadKind(kind){
    return findFolderFile(kind).then(function(path){
      if (!path){
        setHomeImportStatus(kind, "Sem ficheiro na pasta " + DATA_FOLDERS[kind][0] + ".", false);
        return;
      }
      return readFolderFile(path).then(function(rows){
        var parsed = kind === "venda" ? parseMySalesRows(rows) : parsePurchaseRows(rows);
        applyParsed(kind, parsed);
        var d = new Date();
        setHomeImportStatus(kind, "Valores de " + path + " (lido às " + String(d.getHours()).padStart(2,"0") + ":" +
          String(d.getMinutes()).padStart(2,"0") + (parsed.year ? ", meses de " + parsed.year : "") + ").", false);
      });
    }).catch(function(e){
      setHomeImportStatus(kind, "Não foi possível ler o ficheiro de " + KIND_LABEL[kind] + ": " + (e && e.message ? e.message : e) + ".", true);
    });
  }

  // Re-reads both folders (on start and every time "Início" is opened), so edits to the Excel files show up.
  function loadFolderData(){
    if (folderLoad) return folderLoad;
    salesData = {};
    mySales = {};
    folderLoad = Promise.all([loadKind("compra"), loadKind("venda")]).then(function(){
      folderLoad = null;
      if (currentView === "home") renderHome();
    });
    return folderLoad;
  }

  function shiftHomeMonth(delta){
    var parts = homeMonth.split("-");
    var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1 + delta, 1);
    homeMonth = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0");
    renderHome();
  }

  function renderHome(){
    if (!homeMonth) homeMonth = todayStr().slice(0,7);
    var brands = brandList();

    document.getElementById("homeMonthLabel").textContent = monthLabel(homeMonth);
    var monthData = salesData[homeMonth] || {};

    // My own sales value for the month (from the sales Excel).
    var sale = mySales[homeMonth];
    var hasSale = sale !== null && sale !== undefined;
    document.getElementById("homeSaleValue").textContent = hasSale ? formatCurrency(sale) : "—";
    document.getElementById("homeSaleHint").textContent = hasSale
      ? "As minhas vendas em " + monthLabel(homeMonth)
      : "Sem valor de venda para este mês";

    var rows = brands.map(function(b){
      var entry = monthData[findBrandKey(monthData, b)];
      return { brand:b, compra:(entry && entry.compra != null) ? entry.compra : null };
    });
    var totalBuy = 0, brandsWithBuy = 0, maxBuy = 0;
    rows.forEach(function(r){
      if (r.compra === null) return;
      totalBuy += r.compra;
      brandsWithBuy++;
      if (r.compra > maxBuy) maxBuy = r.compra;
    });
    document.getElementById("homePurchaseValue").textContent = brandsWithBuy ? formatCurrency(totalBuy) : "—";
    document.getElementById("homePurchaseHint").textContent = brandsWithBuy
      ? "Soma de " + brandsWithBuy + " marca(s) em " + monthLabel(homeMonth)
      : "Sem valores de compra para este mês";

    // Brands with a value first (largest first), then the ones without a value.
    rows.sort(function(a, b){
      if (a.compra === null && b.compra === null) return 0;
      if (a.compra === null) return 1;
      if (b.compra === null) return -1;
      return b.compra - a.compra;
    });
    var list = document.getElementById("homeBrandList");
    list.innerHTML = "";
    rows.forEach(function(r){
      var li = el("li", "home-brand-row" + (r.compra === null ? " is-empty" : ""));
      var name = el("span", "home-brand-name");
      name.textContent = r.brand;
      var bar = el("span", "home-brand-bar");
      var fill = el("span", "home-brand-bar-fill");
      fill.style.width = (r.compra !== null && maxBuy > 0 ? Math.max(2, Math.round(r.compra / maxBuy * 100)) : 0) + "%";
      bar.appendChild(fill);
      var value = el("span", "home-brand-value mono");
      value.textContent = r.compra !== null ? formatCurrency(r.compra) : "—";
      li.appendChild(name);
      li.appendChild(bar);
      li.appendChild(value);
      list.appendChild(li);
    });
  }

  function render(){
    var isHome = currentView === "home";
    document.getElementById("filters").hidden = isHome;
    document.getElementById("homeView").hidden = !isHome;
    document.getElementById("newItemBtn").hidden = isHome;
    if (isHome){
      document.getElementById("taskQuickAdd").hidden = true;
      document.getElementById("board").hidden = true;
      document.getElementById("meetingsView").hidden = true;
      document.getElementById("projectsView").hidden = true;
      document.getElementById("expensesView").hidden = true;
      renderHome();
      return;
    }
    renderFilters();
    document.getElementById("taskQuickAdd").hidden = currentView !== "tasks";
    document.getElementById("board").hidden = currentView !== "tasks";
    document.getElementById("meetingsView").hidden = currentView !== "meetings";
    document.getElementById("projectsView").hidden = currentView !== "projects";
    document.getElementById("expensesView").hidden = currentView !== "expenses";
    if (currentView === "tasks") renderBoard();
    else if (currentView === "meetings") renderMeetingsList();
    else if (currentView === "projects") renderProjectsList();
    else renderExpensesList();
  }

  function renderFilters(){
    var wrap = document.getElementById("filters");
    wrap.innerHTML = "";
    var map = currentMap();

    var clientIds = Object.keys(clients).sort(function(a,b){ return (clients[a].order||0) - (clients[b].order||0); });

    var allChip = el("button", "chip" + (activeClientFilter===null ? " active" : ""));
    allChip.innerHTML = "Todos <span class=\"count mono\">" + Object.keys(map).length + "</span>";
    allChip.addEventListener("click", function(){ activeClientFilter = null; render(); });
    wrap.appendChild(allChip);

    clientIds.forEach(function(cid){
      var count = 0;
      for (var id in map) if (map[id].clientId === cid) count++;
      var chip = el("button", "chip" + (activeClientFilter===cid ? " active" : ""));
      var dot = el("span", "dot"); dot.style.background = clients[cid].color;
      chip.appendChild(dot);
      chip.appendChild(document.createTextNode(clients[cid].name));
      chip.appendChild(el("span", "count mono", String(count)));
      chip.addEventListener("click", function(){
        activeClientFilter = (activeClientFilter === cid) ? null : cid;
        render();
      });
      wrap.appendChild(chip);
    });

    var manageBtn = el("button", "chip chip-manage", "Gerir clientes");
    manageBtn.addEventListener("click", openClientsModal);
    wrap.appendChild(manageBtn);
  }

  function visibleIds(){
    var map = currentMap();
    return Object.keys(map).filter(function(id){
      return activeClientFilter === null || map[id].clientId === activeClientFilter;
    });
  }

  // ---------- tasks: compact list grouped by status ----------
  var STATUS_NEXT = { todo:"doing", doing:"done", done:"todo" };
  var STATUS_ICON = { todo:"", doing:"", done:"✓" };
  var flashTaskId = null; // row to highlight after it was added or moved
  var doneCollapsed = true;
  try { doneCollapsed = localStorage.getItem("qt_tasks_done_open") !== "1"; } catch(e){}

  // How the list is organised: "status" (A Fazer / Em Curso / Concluído) or "client".
  var taskGrouping = "status";
  var showDoneByClient = false;
  var collapsedClients = {}; // client id ("none" = sem cliente) → true when its group is closed
  try {
    taskGrouping = localStorage.getItem("qt_tasks_grouping") === "client" ? "client" : "status";
    showDoneByClient = localStorage.getItem("qt_tasks_show_done") === "1";
    collapsedClients = JSON.parse(localStorage.getItem("qt_tasks_collapsed_clients") || "{}") || {};
  } catch(e){}
  function saveTaskViewPrefs(){
    try {
      localStorage.setItem("qt_tasks_grouping", taskGrouping);
      localStorage.setItem("qt_tasks_show_done", showDoneByClient ? "1" : "0");
      localStorage.setItem("qt_tasks_collapsed_clients", JSON.stringify(collapsedClients));
    } catch(e){}
  }
  function renderTaskToolbar(){
    Array.prototype.forEach.call(document.querySelectorAll("#taskGroupingSeg button"), function(b){
      b.classList.toggle("active", b.dataset.val === taskGrouping);
      b.setAttribute("aria-pressed", b.dataset.val === taskGrouping ? "true" : "false");
    });
    var showDone = document.getElementById("taskShowDone");
    showDone.checked = showDoneByClient;
    showDone.parentNode.hidden = taskGrouping !== "client";
  }

  function sortedIdsForStatus(statusId, ids){
    var list = ids.filter(function(id){ return tasks[id].status === statusId; });
    if (statusId === "done"){
      // Most recently finished first.
      return list.sort(function(a,b){ return (tasks[b].updatedAt||0) - (tasks[a].updatedAt||0); });
    }
    return list.sort(function(a,b){
      var ta = tasks[a], tb = tasks[b];
      var pa = PRIORITY_ORDER[ta.priority] != null ? PRIORITY_ORDER[ta.priority] : 1;
      var pb = PRIORITY_ORDER[tb.priority] != null ? PRIORITY_ORDER[tb.priority] : 1;
      if (pa !== pb) return pa - pb;
      var da = ta.due || "9999-99-99", db2 = tb.due || "9999-99-99";
      if (da !== db2) return da < db2 ? -1 : 1;
      return (ta.createdAt||0) - (tb.createdAt||0);
    });
  }

  function renderBoard(){
    renderTaskToolbar();
    if (taskGrouping === "client") renderBoardByClient();
    else renderBoardByStatus();
    flashTaskId = null;
  }

  // One group per client (in the client order), "Sem cliente" last. Inside a group:
  // Em Curso first, then A Fazer, then (optionally) Concluído. Drag a row onto another
  // client's group to move it to that client.
  var STATUS_ORDER_IN_CLIENT = { doing:0, todo:1, done:2 };
  function renderBoardByClient(){
    var board = document.getElementById("board");
    board.innerHTML = "";
    var visIds = visibleIds();
    var clientIds = Object.keys(clients).sort(function(a,b){ return (clients[a].order||0) - (clients[b].order||0); });
    clientIds.push(null);

    var shown = 0;
    clientIds.forEach(function(cid){
      var key = cid || "none";
      var all = visIds.filter(function(id){ return (tasks[id].clientId || null) === cid || (cid === null && !clients[tasks[id].clientId]); });
      var open = all.filter(function(id){ return tasks[id].status !== "done"; });
      var doneCount = all.length - open.length;
      var ids = showDoneByClient ? all : open;
      if (!ids.length && !(activeClientFilter && activeClientFilter === cid)) return; // hide clients with nothing to show
      shown++;
      ids = ids.slice().sort(function(a, b){
        var sa = STATUS_ORDER_IN_CLIENT[tasks[a].status], sb = STATUS_ORDER_IN_CLIENT[tasks[b].status];
        if (sa !== sb) return sa - sb;
        var order = sortedIdsForStatus(tasks[a].status, [a, b]);
        return order[0] === a ? -1 : 1;
      });

      var collapsed = !!collapsedClients[key];
      var group = el("section", "task-group task-group-client" + (collapsed ? " is-collapsed" : ""));
      var color = cid ? clientColor(cid) : null;
      if (color) group.style.setProperty("--client-color", color);
      var head = el("button", "task-group-head");
      head.type = "button";
      head.setAttribute("aria-expanded", collapsed ? "false" : "true");
      var dot = el("span", "client-dot");
      dot.style.background = color || "var(--text-faint)";
      head.appendChild(dot);
      var name = el("span", "task-group-title task-group-client-name");
      name.textContent = cid ? clients[cid].name : "Sem cliente";
      head.appendChild(name);
      head.appendChild(el("span", "column-count mono", String(open.length)));
      if (doneCount){
        var doneInfo = el("span", "task-group-done-info");
        doneInfo.textContent = doneCount + (doneCount === 1 ? " concluída" : " concluídas");
        head.appendChild(doneInfo);
      }
      head.appendChild(el("span", "task-group-chevron", collapsed ? "▸" : "▾"));
      head.addEventListener("click", function(){
        if (collapsedClients[key]) delete collapsedClients[key]; else collapsedClients[key] = true;
        saveTaskViewPrefs();
        renderBoard();
      });
      group.appendChild(head);

      group.addEventListener("dragover", function(e){ if (dragTaskId){ e.preventDefault(); group.classList.add("drag-over"); } });
      group.addEventListener("dragleave", function(e){ if (!group.contains(e.relatedTarget)) group.classList.remove("drag-over"); });
      group.addEventListener("drop", function(e){
        e.preventDefault();
        group.classList.remove("drag-over");
        if (dragTaskId) setTaskClient(dragTaskId, cid);
      });

      if (!collapsed){
        var list = el("div", "task-list");
        if (!ids.length) list.appendChild(el("div", "task-empty", "Sem tarefas por fazer."));
        ids.forEach(function(id){ list.appendChild(renderTaskRow(id, true)); });
        group.appendChild(list);
      }
      board.appendChild(group);
    });
    if (!shown){
      board.appendChild(el("div", "task-empty task-empty-all", showDoneByClient ? "Ainda não há tarefas." : "Nada por fazer. Escreve uma tarefa acima."));
    }
  }

  function renderBoardByStatus(){
    var board = document.getElementById("board");
    board.innerHTML = "";
    var visIds = visibleIds();

    STATUSES.forEach(function(status){
      var ids = sortedIdsForStatus(status.id, visIds);
      var collapsible = status.id === "done";
      var collapsed = collapsible && doneCollapsed;

      var group = el("section", "task-group task-group-" + status.id + (collapsed ? " is-collapsed" : ""));
      var head = el(collapsible ? "button" : "div", "task-group-head");
      if (collapsible){
        head.type = "button";
        head.setAttribute("aria-expanded", collapsed ? "false" : "true");
        head.addEventListener("click", function(){
          doneCollapsed = !doneCollapsed;
          try { localStorage.setItem("qt_tasks_done_open", doneCollapsed ? "0" : "1"); } catch(e){}
          renderBoard();
        });
      }
      head.appendChild(el("span", "task-group-title", status.label));
      head.appendChild(el("span", "column-count mono", String(ids.length)));
      if (collapsible) head.appendChild(el("span", "task-group-chevron", collapsed ? "▸" : "▾"));
      group.appendChild(head);

      // Drag a row onto another group to change its status (mouse; on touch use the status button).
      group.addEventListener("dragover", function(e){ if (dragTaskId){ e.preventDefault(); group.classList.add("drag-over"); } });
      group.addEventListener("dragleave", function(e){ if (!group.contains(e.relatedTarget)) group.classList.remove("drag-over"); });
      group.addEventListener("drop", function(e){
        e.preventDefault();
        group.classList.remove("drag-over");
        if (dragTaskId) setTaskStatus(dragTaskId, status.id);
      });

      if (!collapsed){
        var list = el("div", "task-list");
        if (ids.length === 0) list.appendChild(el("div", "task-empty", status.id === "todo" ? "Nada por fazer. Escreve uma tarefa acima." : "Sem tarefas."));
        ids.forEach(function(id){ list.appendChild(renderTaskRow(id)); });
        group.appendChild(list);
      }
      board.appendChild(group);
    });
    flashTaskId = null;
  }

  function renderTaskRow(id, inClientGroup){
    var t = tasks[id];
    var row = el("div", "task-row task-" + t.status + (id === flashTaskId ? " flash" : ""));
    row.draggable = true;
    var color = clientColor(t.clientId);
    if (color) row.style.setProperty("--client-color", color);

    var statusLabel = (STATUSES.filter(function(s){ return s.id === t.status; })[0] || {}).label || "";
    var nextLabel = (STATUSES.filter(function(s){ return s.id === STATUS_NEXT[t.status]; })[0] || {}).label || "";
    var statusBtn = el("button", "task-status status-" + t.status);
    statusBtn.type = "button";
    statusBtn.textContent = STATUS_ICON[t.status] || "";
    statusBtn.title = statusLabel + " — clica para passar a " + nextLabel;
    statusBtn.setAttribute("aria-label", statusBtn.title);
    statusBtn.addEventListener("click", function(e){
      e.stopPropagation();
      setTaskStatus(id, STATUS_NEXT[t.status] || "todo");
    });
    row.appendChild(statusBtn);

    var main = el("div", "task-row-main");
    var title = el("div", "task-title");
    title.textContent = t.title;
    main.appendChild(title);
    var subParts = [];
    var cname = inClientGroup ? null : clientName(t.clientId);
    if (cname) subParts.push(cname);
    if (t.projectId && projects[t.projectId]) subParts.push("Projeto: " + projects[t.projectId].name);
    if (t.points && t.points.length){
      var doneCount = t.points.filter(function(p){ return p.done; }).length;
      subParts.push(doneCount + "/" + t.points.length + " pontos");
    }
    if (t.notes) subParts.push(String(t.notes).split(/\r?\n/)[0]);
    if (subParts.length){
      var sub = el("div", "task-sub");
      if (color && cname){ var dot = el("span", "client-dot"); dot.style.background = color; sub.appendChild(dot); }
      sub.appendChild(document.createTextNode(subParts.join(" · ")));
      main.appendChild(sub);
    }
    row.appendChild(main);

    var side = el("div", "task-side");
    var dueInfo = t.status === "done" ? null : formatDue(t.due);
    if (dueInfo){
      var due = el("span", "due mono " + dueInfo.cls);
      due.textContent = dueInfo.text;
      side.appendChild(due);
    }
    if (t.priority && t.priority !== "media" && t.status !== "done"){
      var prio = el("span", "priority-badge priority-" + t.priority);
      prio.textContent = PRIORITY_LABEL[t.priority];
      side.appendChild(prio);
    }
    row.appendChild(side);

    row.addEventListener("click", function(){ openTaskModal(id); });
    row.addEventListener("dragstart", function(){ dragTaskId = id; row.classList.add("dragging"); });
    row.addEventListener("dragend", function(){ dragTaskId = null; row.classList.remove("dragging"); });
    return row;
  }

  // ---------- quick add: "Ligar ao João amanhã !alta @Cliente" ----------
  var WEEKDAYS = ["domingo","segunda","terca","quarta","quinta","sexta","sabado"];
  function plainWord(s){ return String(s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }
  function dateStr(d){ return d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0") + "-" + String(d.getDate()).padStart(2,"0"); }

  // Returns a due date for a word ("hoje", "amanhã", "sexta-feira", "sábado", "25/11", "25/11/2026"), else null.
  function dueFromWord(word){
    var w = plainWord(word).replace(/[,.;]$/, "");
    var today = new Date(); today.setHours(0,0,0,0);
    if (w === "hoje") return dateStr(today);
    if (w === "amanha"){ today.setDate(today.getDate() + 1); return dateStr(today); }
    var m = w.match(/^(segunda|terca|quarta|quinta|sexta)-feira$/) || w.match(/^(sabado|domingo)$/);
    if (m){
      var target = WEEKDAYS.indexOf(m[1]);
      var diff = (target - today.getDay() + 7) % 7 || 7; // always the next one, never today
      today.setDate(today.getDate() + diff);
      return dateStr(today);
    }
    m = w.match(/^(\d{1,2})[\/.-](\d{1,2})(?:[\/.-](\d{2,4}))?$/);
    if (m){
      var day = parseInt(m[1], 10), month = parseInt(m[2], 10);
      var year = m[3] ? parseInt(m[3], 10) : today.getFullYear();
      if (year < 100) year += 2000;
      var d = new Date(year, month - 1, day);
      if (d.getMonth() !== month - 1 || d.getDate() !== day) return null; // e.g. 31/02
      if (!m[3] && d < today) d.setFullYear(year + 1); // "05/01" in December → next year
      return dateStr(d);
    }
    return null;
  }

  function clientFromWord(word){
    var w = plainWord(word.slice(1));
    if (!w) return null;
    var ids = Object.keys(clients);
    for (var i = 0; i < ids.length; i++){
      if (plainWord(clients[ids[i]].name).replace(/\s+/g, "").indexOf(w) === 0) return ids[i];
    }
    return null;
  }

  function parseQuickTask(text){
    var out = { title:"", priority:"media", due:"", clientId:activeClientFilter || null };
    var rest = [];
    String(text).trim().split(/\s+/).forEach(function(word){
      var w = plainWord(word);
      var p = w.match(/^!(alta|media|baixa|a|m|b)$/);
      if (p){ out.priority = { a:"alta", m:"media", b:"baixa" }[p[1]] || p[1]; return; }
      if (word.charAt(0) === "@" && word.length > 1){
        var cid = clientFromWord(word);
        if (cid){ out.clientId = cid; out.clientExplicit = true; return; }
      }
      if (!out.due){
        var due = dueFromWord(word);
        if (due){ out.due = due; return; }
      }
      rest.push(word);
    });
    out.title = rest.join(" ");
    return out;
  }

  // Shortcuts typed in the title of a NEW task ("Ligar ao João amanhã !alta @cliente"):
  // a live hint under the title, and when leaving the title the prazo/prioridade/cliente
  // fields are filled in and the shortcut words are removed from the title.
  function renderQuickPreview(){
    var preview = document.getElementById("taskQuickPreview");
    var value = document.getElementById("taskTitle").value;
    if (editingTaskId){ preview.hidden = true; return; }
    preview.hidden = false;
    var q = parseQuickTask(value);
    var parts = [];
    if (q.due) parts.push("Prazo: " + (formatDue(q.due) || {}).text);
    if (q.priority !== "media") parts.push("Prioridade: " + PRIORITY_LABEL[q.priority]);
    if (q.clientExplicit) parts.push("Cliente: " + clientName(q.clientId));
    preview.textContent = parts.length
      ? parts.join(" · ")
      : "Atalhos: hoje · amanhã · sexta-feira · 25/11 · !alta · !baixa · @cliente";
    preview.classList.toggle("has-parts", parts.length > 0);
  }

  function applyTitleShortcuts(){
    if (editingTaskId) return;
    var titleInput = document.getElementById("taskTitle");
    var q = parseQuickTask(titleInput.value);
    if (!q.due && q.priority === "media" && !q.clientExplicit) return;
    titleInput.value = q.title;
    if (q.due) document.getElementById("taskDue").value = q.due;
    if (q.priority !== "media") setSegValue("prioritySeg", q.priority);
    if (q.clientExplicit){
      document.getElementById("taskClient").value = q.clientId;
      document.getElementById("taskClientError").hidden = true;
      document.getElementById("taskNewClientRow").hidden = true;
      populateTaskProjectSelect(document.getElementById("taskProject").value);
    }
    renderQuickPreview();
  }

  function openNewTaskFromList(){
    openTaskModal(null, "todo");
  }

  // ---------- meetings list ----------
  function renderMeetingsList(){
    var wrap = document.getElementById("meetingsList");
    wrap.innerHTML = "";
    var ids = visibleIds().sort(function(a,b){
      var da = meetings[a].date || "0000-00-00", db2 = meetings[b].date || "0000-00-00";
      if (da !== db2) return da < db2 ? 1 : -1;
      return (meetings[b].createdAt||0) - (meetings[a].createdAt||0);
    });
    if (ids.length === 0){ wrap.appendChild(el("div", "empty-hint", "Sem reuniões registadas ainda.")); return; }
    ids.forEach(function(id){ wrap.appendChild(renderMeetingRow(id)); });
  }

  function renderMeetingRow(id){
    var m = meetings[id];
    var row = el("div", "meeting-row");
    row.appendChild(el("div", "meeting-row-date mono", formatDateShort(m.date)));
    var main = el("div", "meeting-row-main");
    main.appendChild(el("div", "meeting-row-title", escapeHtml(m.title)));
    var sub = el("div", "meeting-row-sub");
    var color = clientColor(m.clientId);
    if (color){ var dot = el("span", "client-dot"); dot.style.background = color; sub.appendChild(dot); }
    sub.appendChild(document.createTextNode(clientName(m.clientId) || "Sem cliente"));
    main.appendChild(sub);
    row.appendChild(main);
    var pending = (m.actions || []).filter(function(a){ return !a.done; }).length;
    if (pending > 0){
      row.appendChild(el("div", "meeting-row-badge", pending + " ação" + (pending > 1 ? "ões" : "") + " pendente" + (pending > 1 ? "s" : "")));
    }
    row.addEventListener("click", function(){ openMeetingModal(id); });
    return row;
  }

  // ---------- projects list ----------
  function renderProjectsList(){
    var wrap = document.getElementById("projectsList");
    wrap.innerHTML = "";
    var ids = visibleIds().sort(function(a,b){
      var pa = PROJECT_STATUS_ORDER[projects[a].status] != null ? PROJECT_STATUS_ORDER[projects[a].status] : 1;
      var pb = PROJECT_STATUS_ORDER[projects[b].status] != null ? PROJECT_STATUS_ORDER[projects[b].status] : 1;
      if (pa !== pb) return pa - pb;
      return (projects[b].updatedAt||0) - (projects[a].updatedAt||0);
    });
    if (ids.length === 0){ wrap.appendChild(el("div", "empty-hint", "Sem projetos registados ainda.")); return; }
    ids.forEach(function(id){ wrap.appendChild(renderProjectRow(id)); });
  }

  function renderProjectRow(id){
    var p = projects[id];
    var row = el("div", "project-row");
    var main = el("div", "project-row-main");
    main.appendChild(el("div", "project-row-title", escapeHtml(p.name)));
    var sub = el("div", "project-row-sub");
    var color = clientColor(p.clientId);
    if (color){ var dot = el("span", "client-dot"); dot.style.background = color; sub.appendChild(dot); }
    sub.appendChild(document.createTextNode(clientName(p.clientId) || "Sem cliente"));
    main.appendChild(sub);
    row.appendChild(main);

    var docCount = (p.documents || []).length;
    if (docCount > 0) row.appendChild(el("div", "project-row-docs mono", docCount + " doc" + (docCount > 1 ? "s" : "")));

    row.appendChild(el("div", "project-status-badge project-status-" + (p.status || "ativo"), PROJECT_STATUS_LABEL[p.status] || "Ativo"));
    row.addEventListener("click", function(){ openProjectModal(id); });
    return row;
  }

  // ---------- expenses monthly map ----------
  function ensureCurrentMonth(){ if (!currentMonth) currentMonth = todayStr().slice(0,7); }
  function monthLabel(ym){
    var parts = ym.split("-");
    var y = parts[0], m = parseInt(parts[1], 10) - 1;
    return FULL_MONTHS[m] + " " + y;
  }
  function shiftMonth(delta){
    ensureCurrentMonth();
    var parts = currentMonth.split("-");
    var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1;
    var d = new Date(y, m + delta, 1);
    currentMonth = d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0");
    render();
  }
  function idsForCurrentMonth(){
    ensureCurrentMonth();
    return visibleIds().filter(function(id){ return (expenses[id].date || "").slice(0,7) === currentMonth; });
  }

  function renderExpensesList(){
    ensureCurrentMonth();
    document.getElementById("monthLabel").textContent = monthLabel(currentMonth);
    var ids = idsForCurrentMonth();
    var byDay = {};
    ids.forEach(function(id){
      var x = expenses[id];
      var cat = expenseCategory(x);
      var amt = Number(x.amount) || 0;
      if (!byDay[x.date]){
        byDay[x.date] = { total:0 };
        EXPENSE_CATEGORIES.forEach(function(c){ byDay[x.date][c] = 0; });
      }
      byDay[x.date][cat] += amt;
      byDay[x.date].total += amt;
    });

    var table = document.getElementById("expensesTable");
    var days = Object.keys(byDay).sort();

    if (days.length === 0){
      table.innerHTML = "";
      var wrap = table.parentNode;
      if (!wrap.querySelector(".expenses-table-empty")){
        wrap.appendChild(el("div", "expenses-table-empty", "Sem despesas registadas em " + monthLabel(currentMonth) + "."));
      }
      return;
    }
    var oldEmpty = table.parentNode.querySelector(".expenses-table-empty");
    if (oldEmpty) oldEmpty.remove();

    var grand = { total:0 };
    EXPENSE_CATEGORIES.forEach(function(c){ grand[c] = 0; });
    days.forEach(function(d){
      EXPENSE_CATEGORIES.forEach(function(c){ grand[c] += byDay[d][c]; });
      grand.total += byDay[d].total;
    });

    var html = "<thead><tr><th>Dia</th>";
    EXPENSE_CATEGORIES.forEach(function(c){ html += "<th>" + EXPENSE_CATEGORY_LABEL[c] + "</th>"; });
    html += "<th class=\"total-cell\">Total</th></tr></thead><tbody>";
    days.forEach(function(d){
      var row = byDay[d];
      html += "<tr data-date=\"" + d + "\"><td>" + escapeHtml(formatDateShort(d)) + "</td>";
      EXPENSE_CATEGORIES.forEach(function(c){
        html += row[c] > 0 ? "<td class=\"mono\">" + formatCurrency(row[c]) + "</td>" : "<td class=\"mono amount-empty\">—</td>";
      });
      html += "<td class=\"mono total-cell\">" + formatCurrency(row.total) + "</td></tr>";
    });
    html += "</tbody><tfoot><tr><th>Total</th>";
    EXPENSE_CATEGORIES.forEach(function(c){ html += "<td class=\"mono\">" + formatCurrency(grand[c]) + "</td>"; });
    html += "<td class=\"mono total-cell\">" + formatCurrency(grand.total) + "</td></tr></tfoot>";
    table.innerHTML = html;

    Array.prototype.forEach.call(table.querySelectorAll("tbody tr"), function(tr){
      tr.addEventListener("click", function(){ openDayExpensesModal(tr.dataset.date); });
    });
  }

  function renderDayExpensesList(date){
    document.getElementById("dayExpensesTitle").textContent = "Despesas de " + formatDateShort(date);
    var list = document.getElementById("dayExpensesList");
    list.innerHTML = "";
    var ids = visibleIds().filter(function(id){ return expenses[id].date === date; })
      .sort(function(a,b){ return (expenses[a].createdAt||0) - (expenses[b].createdAt||0); });
    if (ids.length === 0){ list.appendChild(el("div", "hint", "Sem despesas neste dia.")); return; }
    ids.forEach(function(id){
      var x = expenses[id];
      var row = el("div", "day-expense-item");
      var cat = expenseCategory(x);
      var badge = el("div", "expense-cat-badge expense-cat-" + cat);
      badge.textContent = EXPENSE_CATEGORY_LABEL[cat] + (cat === "combustivel" && x.payment === "frota" ? " · frota" : "");
      row.appendChild(badge);
      var descWrap = el("div", "day-expense-desc");
      var companion = mealCompanion(x);
      descWrap.textContent = (x.description || "") + (companion ? (x.description ? " · " : "") + (companion === "sozinho" ? "sozinho" : "com " + companion) : "");
      if (x.supplier){
        var supplier = el("span", "hint-inline");
        supplier.textContent = " — " + x.supplier;
        descWrap.appendChild(supplier);
      }
      row.appendChild(descWrap);
      row.appendChild(el("div", "day-expense-amount mono", formatCurrency(x.amount)));
      row.addEventListener("click", function(){ closeDayExpensesModal(); openExpenseModal(id); });
      list.appendChild(row);
    });
  }

  function openDayExpensesModal(date){
    dayModalDate = date;
    renderDayExpensesList(date);
    document.getElementById("dayExpensesModalOverlay").hidden = false;
  }
  function closeDayExpensesModal(){
    document.getElementById("dayExpensesModalOverlay").hidden = true;
    dayModalDate = null;
  }

  function exportExpensesCSV(){
    var statusEl = document.getElementById("expensesExportStatus");
    statusEl.classList.remove("error");
    var ids = idsForCurrentMonth().sort(function(a,b){
      var da = expenses[a].date || "0000-00-00", db2 = expenses[b].date || "0000-00-00";
      return da < db2 ? -1 : (da > db2 ? 1 : 0);
    });
    function csvField(v){
      var s = String(v == null ? "" : v);
      if (/[;"\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
      return s;
    }
    var lines = ["Data;Categoria;Com quem;Cliente;Fornecedor;Descrição;Valor (€)"];
    ids.forEach(function(id){
      var x = expenses[id];
      lines.push([
        csvField(x.date || ""),
        csvField(EXPENSE_CATEGORY_LABEL[expenseCategory(x)] + (x.payment === "frota" ? " (cartão frota)" : "")),
        csvField(mealCompanion(x)),
        csvField(clientName(x.clientId) || ""),
        csvField(x.supplier || ""),
        csvField(x.description || ""),
        csvField((Number(x.amount)||0).toFixed(2).replace(".", ","))
      ].join(";"));
    });
    var csv = lines.join("\r\n");
    try {
      var blob = new Blob([csv], { type:"text/csv;charset=utf-8" });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = "despesas-" + currentMonth + ".csv";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
      statusEl.textContent = "Descarregado.";
      setTimeout(function(){ statusEl.textContent = ""; }, 3000);
    } catch(err){
      console.error("export csv failed", err);
      statusEl.textContent = "Não foi possível exportar.";
      statusEl.classList.add("error");
    }
  }

  // ---------- task ops ----------
  function setTaskClient(id, clientId){
    if (!tasks[id] || (tasks[id].clientId || null) === (clientId || null)) return;
    tasks[id].clientId = clientId || null; tasks[id].updatedAt = Date.now();
    flashTaskId = id;
    saveStore(); render();
  }
  function setTaskStatus(id, statusId){
    if (!tasks[id] || tasks[id].status === statusId) return;
    tasks[id].status = statusId; tasks[id].updatedAt = Date.now();
    flashTaskId = id;
    saveStore(); render();
  }
  function saveTaskFromModal(){
    applyTitleShortcuts();
    var title = document.getElementById("taskTitle").value.trim();
    if (!title){ document.getElementById("taskTitle").focus(); return; }
    var clientId = document.getElementById("taskClient").value;
    if (clientId === "__new__") clientId = "";
    // New tasks must have a client; older tasks without one can still be edited.
    if (!clientId && (!editingTaskId || tasks[editingTaskId].clientId)){
      document.getElementById("taskClientError").hidden = false;
      document.getElementById("taskClient").focus();
      return;
    }
    // A point typed but not yet added with Enter is kept too.
    if (document.getElementById("newPointInput").value.trim()) addPointDraft();
    var payload = {
      title:title,
      clientId: clientId || null,
      projectId: document.getElementById("taskProject").value || null,
      points: pointsDraft.slice(),
      priority: document.getElementById("prioritySeg").dataset.value,
      status: document.getElementById("statusSeg").dataset.value,
      due: document.getElementById("taskDue").value || "",
      notes: document.getElementById("taskNotes").value.trim(),
      updatedAt: Date.now()
    };
    if (editingTaskId){ tasks[editingTaskId] = Object.assign({}, tasks[editingTaskId], payload); }
    else { payload.createdAt = Date.now(); var newId = genId(); tasks[newId] = payload; flashTaskId = newId; }
    saveStore(); closeTaskModal(); render();
  }
  function deleteEditingTask(){
    if (!editingTaskId) return;
    delete tasks[editingTaskId];
    saveStore(); closeTaskModal(); render();
  }

  // ---------- meeting ops ----------
  function saveMeetingFromModal(){
    var title = document.getElementById("meetingTitle").value.trim();
    if (!title){ document.getElementById("meetingTitle").focus(); return; }
    var payload = {
      title:title,
      clientId: document.getElementById("meetingClient").value || null,
      date: document.getElementById("meetingDate").value || todayStr(),
      agenda: document.getElementById("meetingAgenda").value.trim(),
      notes: document.getElementById("meetingNotes").value.trim(),
      summary: document.getElementById("meetingSummary").value.trim(),
      actions: actionsDraft.map(function(a){ return { id:a.id, text:a.text, done:!!a.done }; }),
      updatedAt: Date.now()
    };
    if (editingMeetingId){ meetings[editingMeetingId] = Object.assign({}, meetings[editingMeetingId], payload); }
    else { payload.createdAt = Date.now(); meetings[genId()] = payload; }
    saveStore(); closeMeetingModal(); render();
  }
  function deleteEditingMeeting(){
    if (!editingMeetingId) return;
    delete meetings[editingMeetingId];
    saveStore(); closeMeetingModal(); render();
  }

  function renderActionsDraft(){
    var list = document.getElementById("actionsDraftList");
    list.innerHTML = "";
    if (actionsDraft.length === 0){ list.appendChild(el("div", "hint", "Sem ações associadas.")); return; }
    actionsDraft.forEach(function(item){
      var row = el("div", "action-item");
      var cb = document.createElement("input");
      cb.type = "checkbox"; cb.checked = !!item.done;
      cb.addEventListener("change", function(){ item.done = cb.checked; renderActionsDraft(); });
      row.appendChild(cb);
      row.appendChild(el("span", "action-text" + (item.done ? " done" : ""), escapeHtml(item.text)));
      var rm = el("button", "icon-btn", "×");
      rm.type = "button"; rm.title = "Remover";
      rm.addEventListener("click", function(){
        actionsDraft = actionsDraft.filter(function(a){ return a.id !== item.id; });
        renderActionsDraft();
      });
      row.appendChild(rm);
      list.appendChild(row);
    });
  }
  function addActionDraft(text){
    if (!text.trim()) return;
    actionsDraft.push({ id:genId(), text:text.trim(), done:false });
    renderActionsDraft();
  }

  // ---------- project ops ----------
  function saveProjectFromModal(){
    var name = document.getElementById("projectName").value.trim();
    if (!name){ document.getElementById("projectName").focus(); return; }
    var payload = {
      name:name,
      clientId: document.getElementById("projectClient").value || null,
      status: document.getElementById("projectStatusSeg").dataset.value,
      description: document.getElementById("projectDescription").value.trim(),
      documents: documentsDraft.map(function(d){ return { id:d.id, name:d.name, contentType:d.contentType, sizeBytes:d.sizeBytes, uploadedAt:d.uploadedAt }; }),
      updatedAt: Date.now()
    };
    if (editingProjectId){ projects[editingProjectId] = Object.assign({}, projects[editingProjectId], payload); }
    else { payload.createdAt = Date.now(); projects[genId()] = payload; }
    documentsPendingCleanup = []; // now referenced by a saved project; keep the blobs
    saveStore(); closeProjectModal(); render();
  }
  function deleteEditingProject(){
    if (!editingProjectId) return;
    var proj = projects[editingProjectId];
    (proj && proj.documents || []).forEach(function(d){ idbDelete(d.id).catch(function(){}); });
    delete projects[editingProjectId];
    saveStore(); closeProjectModal(); render();
  }

  function renderDocumentsDraft(){
    var list = document.getElementById("documentsDraftList");
    list.innerHTML = "";
    if (documentsDraft.length === 0){ list.appendChild(el("div", "hint", "Sem documentos anexados.")); return; }
    documentsDraft.forEach(function(doc){
      var row = el("div", "document-item");
      var ext = (doc.name.split(".").pop() || "?").slice(0,4).toUpperCase();
      row.appendChild(el("div", "document-icon", escapeHtml(ext)));

      var info = el("div", "document-info");
      var nameBtn = el("button", "document-name", escapeHtml(doc.name));
      nameBtn.type = "button";
      nameBtn.addEventListener("click", function(){ openDocumentBlob(doc.id, doc.name); });
      info.appendChild(nameBtn);
      info.appendChild(el("div", "document-meta mono", formatBytes(doc.sizeBytes)));
      row.appendChild(info);

      var rm = el("button", "icon-btn", "×");
      rm.type = "button"; rm.title = "Remover documento";
      rm.addEventListener("click", function(){ removeDocumentDraft(doc.id); });
      row.appendChild(rm);
      list.appendChild(row);
    });
  }

  function openDocumentBlob(id, name){
    idbGet(id).then(function(rec){
      if (!rec || !rec.blob){ alert("Não foi possível encontrar este documento neste navegador."); return; }
      var url = URL.createObjectURL(rec.blob);
      window.open(url, "_blank");
      setTimeout(function(){ URL.revokeObjectURL(url); }, 60000);
    }).catch(function(err){
      console.error("openDocumentBlob failed", err);
      alert("Não foi possível abrir \"" + name + "\".");
    });
  }

  function removeDocumentDraft(docId){
    documentsDraft = documentsDraft.filter(function(d){ return d.id !== docId; });
    documentsPendingCleanup = documentsPendingCleanup.filter(function(id){ return id !== docId; });
    renderDocumentsDraft();
    idbDelete(docId).catch(function(err){ console.error("idbDelete failed", err); });
  }

  function addDocumentsFromFiles(fileList){
    var statusEl = document.getElementById("documentAddStatus");
    var files = Array.prototype.slice.call(fileList || []);
    if (files.length === 0) return;
    statusEl.classList.remove("error");
    var remaining = files.length;
    statusEl.textContent = "A anexar " + remaining + " ficheiro" + (remaining > 1 ? "s" : "") + "...";

    files.forEach(function(file){
      var id = genId();
      idbPut({ id:id, name:file.name, contentType:file.type, sizeBytes:file.size, blob:file }).then(function(){
        documentsDraft.push({ id:id, name:file.name, contentType:file.type, sizeBytes:file.size, uploadedAt:Date.now() });
        documentsPendingCleanup.push(id);
        renderDocumentsDraft();
      }).catch(function(err){
        console.error("idbPut failed", err);
        statusEl.textContent = "Não foi possível anexar \"" + file.name + "\".";
        statusEl.classList.add("error");
      }).finally(function(){
        remaining -= 1;
        if (remaining === 0 && !statusEl.classList.contains("error")) statusEl.textContent = "";
      });
    });
  }

  // ---------- expense ops ----------
  function saveExpenseFromModal(){
    var amount = parseAmountInput(document.getElementById("expenseAmount").value);
    var description = document.getElementById("expenseDescription").value.trim();
    // The value is always required (fleet card included).
    if (amount <= 0){
      document.getElementById("expenseAmountError").hidden = false;
      document.getElementById("expenseAmount").focus();
      return;
    }
    if (!validateMealFields()) return;
    var payload = {
      date: document.getElementById("expenseDate").value || todayStr(),
      category: document.getElementById("expenseCategorySeg").dataset.value,
      payment: null,
      fuel: null,
      mealWith: null,
      mealSupplier: "",
      mealColleague: "",
      clientId: document.getElementById("expenseClient").value || null,
      supplier: document.getElementById("expenseSupplier").value.trim(),
      description: description,
      amount: amount,
      updatedAt: Date.now()
    };
    if (isMealCategory(payload.category)){
      payload.mealWith = document.getElementById("mealWithSeg").dataset.value;
      if (payload.mealWith !== "cliente") payload.clientId = null;
      if (payload.mealWith === "fornecedor") payload.mealSupplier = document.getElementById("mealSupplier").value.trim();
      if (payload.mealWith === "colega") payload.mealColleague = document.getElementById("mealSupplier").value.trim();
    } else {
      // Only meals have a client / supplier (restaurant); other categories never do.
      payload.clientId = null;
      payload.supplier = "";
    }
    if (payload.category === "combustivel"){
      // The fleet-card station is shown as the supplier in the day list.
      payload.payment = document.getElementById("expensePaymentSeg").dataset.value;
      if (payload.payment === "frota"){
        payload.fuel = readFuelFields();
        payload.supplier = payload.fuel.station;
      }
    }
    if (editingExpenseId){ expenses[editingExpenseId] = Object.assign({}, expenses[editingExpenseId], payload); }
    else { payload.createdAt = Date.now(); expenses[genId()] = payload; }
    saveStore(); closeExpenseModal(); render();
  }
  function deleteEditingExpense(){
    if (!editingExpenseId) return;
    delete expenses[editingExpenseId];
    saveStore(); closeExpenseModal(); render();
  }

  // ---------- meals (almoços / jantares / refeições gerais): with whom ----------
  // "Sozinho", "Cliente" (client required), "Fornecedor" or "Colega" (name required).
  // The supplier and colleague names share one input; each keeps its own value while switching.
  var MEAL_NAME_FIELD = {
    fornecedor: { key:"mealSupplier", label:"Qual fornecedor", placeholder:"Ex: Kennametal", error:"Indica o fornecedor.", base:DEFAULT_BRANDS },
    colega: { key:"mealColleague", label:"Qual colega", placeholder:"Ex: Nuno Teixeira", error:"Indica o colega.", base:[] }
  };
  var mealNameDraft = { fornecedor:"", colega:"" };
  var mealNameShown = null; // which kind the shared input currently holds

  function isMealCategory(cat){ return cat === "almocos" || cat === "jantares" || cat === "refeicoes"; }
  function mealCompanion(x){
    if (!x || !isMealCategory(expenseCategory(x)) || !x.mealWith) return "";
    if (x.mealWith === "cliente") return clientName(x.clientId) || "cliente";
    if (x.mealWith === "fornecedor") return x.mealSupplier || "fornecedor";
    if (x.mealWith === "colega") return x.mealColleague || "colega";
    return "sozinho";
  }
  function fillMealNameSuggestions(kind){
    var cfg = MEAL_NAME_FIELD[kind];
    var seen = {}, names = [];
    cfg.base.concat(Object.keys(expenses).map(function(id){ return expenses[id][cfg.key] || ""; })).forEach(function(n){
      var k = String(n).trim().toLowerCase();
      if (k && !seen[k]){ seen[k] = true; names.push(String(n).trim()); }
    });
    var list = document.getElementById("mealSupplierList");
    list.innerHTML = "";
    names.sort(function(a, b){ return a.localeCompare(b, "pt"); }).forEach(function(n){
      var o = document.createElement("option"); o.value = n; list.appendChild(o);
    });
  }
  function fillMealFields(x){
    setSegValue("mealWithSeg", (x && x.mealWith) || "");
    mealNameDraft = { fornecedor:(x && x.mealSupplier) || "", colega:(x && x.mealColleague) || "" };
    mealNameShown = null;
    ["mealWithError","mealSupplierError","expenseClientError"].forEach(function(id){ document.getElementById(id).hidden = true; });
    updateMealFields();
  }
  function updateMealFields(){
    var cat = document.getElementById("expenseCategorySeg").dataset.value;
    var meal = isMealCategory(cat);
    var withWhom = document.getElementById("mealWithSeg").dataset.value;
    var input = document.getElementById("mealSupplier");
    // Keep what was typed for the kind shown before, then show the chosen kind.
    if (mealNameShown) mealNameDraft[mealNameShown] = input.value;
    var nameKind = meal && MEAL_NAME_FIELD[withWhom] ? withWhom : null;
    if (nameKind && nameKind !== mealNameShown){
      var cfg = MEAL_NAME_FIELD[nameKind];
      input.value = mealNameDraft[nameKind];
      input.placeholder = cfg.placeholder;
      document.getElementById("mealSupplierLabel").textContent = cfg.label;
      document.getElementById("mealSupplierError").textContent = cfg.error;
      document.getElementById("mealSupplierError").hidden = true;
      fillMealNameSuggestions(nameKind);
    }
    mealNameShown = nameKind;
    document.getElementById("mealFields").hidden = !meal;
    document.getElementById("mealSupplierField").hidden = !nameKind;
    // Client / restaurant only exist for meals.
    document.getElementById("expenseClientSupplierRow").hidden = !meal;
    // In meals the client field only shows (and is required) when eating with a client.
    document.getElementById("expenseClientField").hidden = meal && withWhom !== "cliente";
    document.getElementById("expenseClientOptional").hidden = meal && withWhom === "cliente";
    document.getElementById("expenseSupplierLabel").textContent = meal ? "Restaurante" : "Fornecedor";
    document.getElementById("expenseSupplier").placeholder = meal ? "Ex: Restaurante O Lavrador" : "Ex: Galp, Continente...";
    if (withWhom) document.getElementById("mealWithError").hidden = true;
  }
  // Returns false (and shows the message) when the meal's company is missing.
  function validateMealFields(){
    var cat = document.getElementById("expenseCategorySeg").dataset.value;
    if (!isMealCategory(cat)) return true;
    var withWhom = document.getElementById("mealWithSeg").dataset.value;
    if (!withWhom){
      document.getElementById("mealWithError").hidden = false;
      document.getElementById("mealWithSeg").scrollIntoView({ block:"center" });
      return false;
    }
    if (withWhom === "cliente" && !document.getElementById("expenseClient").value){
      document.getElementById("expenseClientError").hidden = false;
      document.getElementById("expenseClient").focus();
      return false;
    }
    if (MEAL_NAME_FIELD[withWhom] && !document.getElementById("mealSupplier").value.trim()){
      document.getElementById("mealSupplierError").hidden = false;
      document.getElementById("mealSupplier").focus();
      return false;
    }
    return true;
  }

  // ---------- fuel (fleet card) fields in the expense window ----------
  // Card number, name and product are remembered on this device for the next fill-in.
  var FUEL_DEFAULTS_KEY = "qt_fuel_defaults";
  function fuelDefaults(){
    try { return JSON.parse(localStorage.getItem(FUEL_DEFAULTS_KEY) || "{}") || {}; } catch(e){ return {}; }
  }
  function fillFuelFields(x){
    var f = (x && x.fuel) || {};
    var d = x ? {} : fuelDefaults();
    setSegValue("expensePaymentSeg", (x && x.payment) || "frota");
    document.getElementById("fuelCard").value = f.card || d.card || "";
    document.getElementById("fuelTime").value = f.time || "";
    document.getElementById("fuelStation").value = f.station || "";
    document.getElementById("fuelZone").value = f.zone || "";
    document.getElementById("fuelName").value = f.name || d.name || "";
    document.getElementById("fuelKms").value = f.kms != null ? String(f.kms) : "";
    setFuelProduct(f.product || d.product || "Gasóleo");
    document.getElementById("fuelLiters").value = f.liters != null ? String(f.liters).replace(".", ",") : "";
    updateFuelFields();
  }
  function setFuelProduct(value){
    var sel = document.getElementById("fuelProduct");
    var exists = Array.prototype.some.call(sel.options, function(o){ return o.value === value; });
    if (!exists){
      var o = document.createElement("option");
      o.value = value; o.textContent = value;
      sel.appendChild(o);
    }
    sel.value = value;
  }
  function readFuelFields(){
    function val(id){ return document.getElementById(id).value.trim(); }
    var kms = parseInt(val("fuelKms").replace(/\D/g, ""), 10);
    var liters = parseAmountInput(val("fuelLiters"));
    var f = {
      card: val("fuelCard"), time: val("fuelTime"), station: val("fuelStation"), zone: val("fuelZone"),
      name: val("fuelName"), kms: isNaN(kms) ? null : kms, product: val("fuelProduct") || "Gasóleo",
      liters: liters > 0 ? liters : null
    };
    try { localStorage.setItem(FUEL_DEFAULTS_KEY, JSON.stringify({ card:f.card, name:f.name, product:f.product })); } catch(e){}
    return f;
  }
  function updateFuelFields(){
    var isFuel = document.getElementById("expenseCategorySeg").dataset.value === "combustivel";
    var fleet = document.getElementById("expensePaymentSeg").dataset.value === "frota";
    document.getElementById("expenseFuelFields").hidden = !isFuel;
    document.getElementById("expenseFleetFields").hidden = !fleet;
    document.getElementById("expensePaymentHint").textContent = fleet
      ? "Vai para o Excel Combustiveis do mês."
      : "Vai para a coluna Gasóleo do mapa de Despesas.";
  }

  // ---------- automatic Excel files (only when opened through servidor_local.py) ----------
  // From the templates in the "Despesas" folder the app writes:
  //   "Despesas AAAA.xlsx"        — everything except fleet-card fuel, in the template's columns
  //   "Combustiveis AAAA-MM.xlsx" — fleet-card fuel of that month
  // A file is only rewritten when its expenses changed (fingerprint kept on this device).
  var EXCEL_COLUMN = { viatura:"D", material:"E", combustivel:"F", almocos:"G", jantares:"H", refeicoes:"I", estadias:"J", parques:"L", portagens:"M", outros:"N" };
  var excelWriterAvailable = null; // null = not checked yet
  var excelSyncTimer = null, excelSyncRunning = false, excelSyncAgain = false;
  var EXCEL_SIG_KEY = "qt_excel_sig";
  var excelFilesSeen = {}; // output files confirmed to exist in the folder during this session

  function isFleetFuel(x){ return expenseCategory(x) === "combustivel" && x.payment === "frota"; }

  function setExcelStatus(text, isError, showRetry){
    var s = document.getElementById("excelSyncStatus");
    s.hidden = !text;
    s.textContent = text || "";
    s.classList.toggle("error", !!isError);
    document.getElementById("excelSyncRetryBtn").hidden = !showRetry;
  }

  function checkExcelWriter(){
    if (excelWriterAvailable !== null) return Promise.resolve(excelWriterAvailable);
    if (!/^(127\.0\.0\.1|localhost)$/.test(location.hostname) || !window.QuadroExcel){
      excelWriterAvailable = false;
      return Promise.resolve(false);
    }
    return fetch("api/estado", { cache:"no-store" }).then(function(r){
      return r.ok ? r.json() : null;
    }).then(function(j){
      excelWriterAvailable = !!(j && j.gravarExcel);
      return excelWriterAvailable;
    }).catch(function(){ excelWriterAvailable = false; return false; });
  }

  function scheduleExcelSync(){
    clearTimeout(excelSyncTimer);
    excelSyncTimer = setTimeout(function(){ runExcelSync(false); }, 1500);
  }

  // Groups expenses into the files they belong to: { "Despesas 2026.xlsx": {...}, "Combustiveis 2026-11.xlsx": {...} }
  function excelOutputs(){
    var out = {};
    function sortKey(x){ return (x.date || "") + " " + ((x.fuel && x.fuel.time) || "") + " " + (x.createdAt || 0); }
    var ids = Object.keys(expenses).sort(function(a, b){
      var ka = sortKey(expenses[a]), kb = sortKey(expenses[b]);
      return ka < kb ? -1 : (ka > kb ? 1 : 0);
    });
    ids.forEach(function(id){
      var x = expenses[id];
      if (!/^\d{4}-\d{2}-\d{2}$/.test(x.date || "")) return;
      var amount = Math.round((Number(x.amount) || 0) * 100) / 100;
      if (isFleetFuel(x)){
        var name = "Combustiveis " + x.date.slice(0, 7) + ".xlsx";
        out[name] = out[name] || { kind:"combustiveis", rows:[] };
        var f = x.fuel || {};
        out[name].rows.push({
          card:f.card || "", date:x.date, time:f.time || "", station:f.station || x.supplier || "", zone:f.zone || "",
          name:f.name || "", kms:f.kms, price:amount > 0 ? amount : null, product:f.product || "Gasóleo", liters:f.liters
        });
      } else {
        var year = x.date.slice(0, 4);
        var dname = "Despesas " + year + ".xlsx";
        out[dname] = out[dname] || { kind:"despesas", year:+year, items:[] };
        var cat = expenseCategory(x);
        var note = cat === "estadias" ? (x.supplier || x.description || "") : "";
        // Comment on the day/category cell: the description, plus who the meal was with.
        var companion = mealCompanion(x);
        var commentParts = [];
        if (x.description) commentParts.push(x.description);
        if (companion) commentParts.push(companion === "sozinho" ? "Sozinho" : "Com " + companion);
        if (isMealCategory(cat) && x.supplier) commentParts.push(x.supplier);
        out[dname].items.push({ date:x.date, col:EXCEL_COLUMN[cat], amount:amount, note:note, comment:commentParts.join(" — ") });
      }
    });
    return out;
  }

  function fetchTemplate(name){
    return fetch("Despesas/" + encodeURIComponent(name), { cache:"no-store" }).then(function(r){
      if (!r.ok) throw new Error("falta o modelo " + name + " na pasta Despesas");
      return r.arrayBuffer();
    });
  }
  function putExcel(name, bytes){
    return fetch("Despesas/" + encodeURIComponent(name), {
      method:"PUT", body:bytes,
      headers:{ "Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
    }).then(function(r){
      if (r.ok) return;
      return r.json().catch(function(){ return {}; }).then(function(j){
        throw new Error(j.erro || ("erro " + r.status + " ao gravar " + name));
      });
    });
  }

  function runExcelSync(force){
    if (excelSyncRunning){ excelSyncAgain = true; return Promise.resolve(); }
    excelSyncRunning = true;
    return checkExcelWriter().then(function(ok){
      if (!ok) return;
      var sigKey = EXCEL_SIG_KEY + ":" + STORAGE_KEY;
      var sigs = {};
      try { sigs = JSON.parse(localStorage.getItem(sigKey) || "{}") || {}; } catch(e){}
      var outputs = excelOutputs();
      // Files written before whose expenses were all deleted are rewritten empty.
      Object.keys(sigs).forEach(function(name){
        if (outputs[name]) return;
        var m = /^Despesas (\d{4})\.xlsx$/.exec(name);
        outputs[name] = m ? { kind:"despesas", year:+m[1], items:[] } : { kind:"combustiveis", rows:[] };
      });
      var templates = {};
      function template(name){
        if (!templates[name]) templates[name] = fetchTemplate(name);
        return templates[name];
      }
      var written = [], errors = [];
      var names = Object.keys(outputs).sort();
      var chain = Promise.resolve();
      names.forEach(function(name){
        var o = outputs[name];
        chain = chain.then(function(){
          var tplName = o.kind === "despesas" ? "Despesas.xlsx" : "Combustiveis.xlsx";
          var sig = JSON.stringify([tplName, o.items || o.rows]);
          // Unchanged expenses: skip without downloading the template, unless the file was
          // deleted from the folder (checked once per session, or when "Gravar Excel agora" is used).
          var needed;
          if (force || sigs[name] !== sig) needed = Promise.resolve(true);
          else if (excelFilesSeen[name]) needed = Promise.resolve(false);
          else needed = fetch("Despesas/" + encodeURIComponent(name), { method:"HEAD", cache:"no-store" })
            .then(function(r){ if (r.ok) excelFilesSeen[name] = true; return !r.ok; })
            .catch(function(){ return false; });
          return needed.then(function(yes){
            if (!yes) return;
            return template(tplName).then(function(tpl){
              return o.kind === "despesas"
                ? window.QuadroExcel.buildDespesas(tpl, o.year, o.items)
                : window.QuadroExcel.buildCombustiveis(tpl, o.rows);
            }).then(function(bytes){ return putExcel(name, bytes); }).then(function(){
              sigs[name] = sig;
              excelFilesSeen[name] = true;
              written.push(name);
            });
          }).catch(function(e){ errors.push(e && e.message ? e.message : String(e)); });
        });
      });
      return chain.then(function(){
        try { localStorage.setItem(sigKey, JSON.stringify(sigs)); } catch(e){}
        var d = new Date(), at = String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
        if (errors.length) setExcelStatus("Excel não atualizado: " + errors[0], true, true);
        else if (written.length) setExcelStatus("Excel atualizado às " + at + ": " + written.join(", "), false, false);
        else if (names.length) setExcelStatus("Excel em dia (pasta Despesas).", false, false);
        else setExcelStatus("", false, false);
      });
    }).then(function(){
      excelSyncRunning = false;
      if (excelSyncAgain){ excelSyncAgain = false; return runExcelSync(false); }
    }, function(e){
      excelSyncRunning = false;
      setExcelStatus("Excel não atualizado: " + (e && e.message ? e.message : e), true, true);
    });
  }

  // ---------- client ops ----------
  function addClient(name){
    var id = genId();
    clients[id] = { name:name.trim(), color:nextPaletteColor(), order:Date.now(), updatedAt:Date.now() };
    saveStore(); renderClientsList(); render();
    return id;
  }
  function renameClient(id, name){
    if (!clients[id] || !name.trim()) return;
    clients[id].name = name.trim();
    clients[id].updatedAt = Date.now();
    saveStore(); render();
  }
  function recolorClient(id, color){
    if (!clients[id]) return;
    clients[id].color = color;
    clients[id].updatedAt = Date.now();
    saveStore(); renderClientsList(); render();
  }
  function deleteClient(id){
    delete clients[id];
    if (activeClientFilter === id) activeClientFilter = null;
    saveStore();
  }

  // ---------- shared client select ----------
  function populateClientSelect(selectId, selectedId){
    var sel = document.getElementById(selectId);
    sel.innerHTML = "";
    var noneOpt = el("option", "", "Sem cliente"); noneOpt.value = "";
    sel.appendChild(noneOpt);
    Object.keys(clients).sort(function(a,b){ return (clients[a].order||0) - (clients[b].order||0); }).forEach(function(cid){
      var opt = el("option", "", escapeHtml(clients[cid].name));
      opt.value = cid;
      sel.appendChild(opt);
    });
    sel.value = selectedId || "";
  }
  function setSegValue(segId, value){
    var seg = document.getElementById(segId);
    seg.dataset.value = value;
    Array.prototype.forEach.call(seg.querySelectorAll("button"), function(btn){
      btn.classList.toggle("active", btn.dataset.val === value);
    });
  }

  // ---------- task modal ----------
  // preset (optional, for new tasks): { title, priority, due, clientId } — e.g. from the quick-add box.
  function openTaskModal(id, presetStatus, preset){
    editingTaskId = id || null;
    var overlay = document.getElementById("taskModalOverlay");
    var title = document.getElementById("taskModalTitle");
    var deleteBtn = document.getElementById("deleteTaskBtn");
    var focusId = "taskTitle";
    if (id && tasks[id]){
      var t = tasks[id];
      title.textContent = "Editar tarefa";
      document.getElementById("taskTitle").value = t.title;
      populateTaskClientSelect(t.clientId);
      populateTaskProjectSelect(t.projectId);
      setSegValue("prioritySeg", t.priority || "media");
      setSegValue("statusSeg", t.status || "todo");
      document.getElementById("taskDue").value = t.due || "";
      document.getElementById("taskNotes").value = t.notes || "";
      pointsDraft = (t.points || []).map(function(p){ return { id:p.id, text:p.text, done:!!p.done }; });
      deleteBtn.hidden = false;
      resetConfirm(deleteBtn, "Eliminar");
    } else {
      var p = preset || {};
      title.textContent = "Nova tarefa";
      document.getElementById("taskTitle").value = p.title || "";
      populateTaskClientSelect(p.clientId || activeClientFilter);
      populateTaskProjectSelect(null);
      setSegValue("prioritySeg", p.priority || "media");
      setSegValue("statusSeg", presetStatus || "todo");
      document.getElementById("taskDue").value = p.due || "";
      document.getElementById("taskNotes").value = "";
      pointsDraft = [];
      deleteBtn.hidden = true;
      // Title already written in the quick box → go straight to the client, then the work points.
      if (p.title) focusId = document.getElementById("taskClient").value ? "newPointInput" : "taskClient";
    }
    document.getElementById("taskClientError").hidden = true;
    document.getElementById("taskNewClientRow").hidden = true;
    document.getElementById("newPointInput").value = "";
    renderPointsDraft();
    renderQuickPreview();
    overlay.hidden = false;
    setTimeout(function(){ document.getElementById(focusId).focus(); }, 0);
  }

  // Client list for the task window, with an option to create a client without leaving it.
  function populateTaskClientSelect(selectedId){
    populateClientSelect("taskClient", selectedId);
    var sel = document.getElementById("taskClient");
    sel.options[0].textContent = "— Escolher cliente —";
    var newOpt = el("option", "", "+ Novo cliente…");
    newOpt.value = "__new__";
    sel.appendChild(newOpt);
  }
  // Existing projects: the chosen client's first, then the others; finished ones marked.
  function populateTaskProjectSelect(selectedId){
    var sel = document.getElementById("taskProject");
    var clientId = document.getElementById("taskClient").value;
    sel.innerHTML = "";
    var none = el("option", "", "Sem projeto");
    none.value = "";
    sel.appendChild(none);
    var ids = Object.keys(projects).sort(function(a, b){
      var sa = PROJECT_STATUS_ORDER[projects[a].status] || 0, sb = PROJECT_STATUS_ORDER[projects[b].status] || 0;
      if (sa !== sb) return sa - sb;
      return String(projects[a].name).localeCompare(String(projects[b].name), "pt");
    });
    function option(pid){
      var p = projects[pid];
      var label = p.name;
      if (p.clientId !== clientId && clientName(p.clientId)) label += " — " + clientName(p.clientId);
      if (p.status === "concluido") label += " (concluído)";
      var o = el("option", "");
      o.textContent = label;
      o.value = pid;
      return o;
    }
    var mine = clientId && clients[clientId] ? ids.filter(function(pid){ return projects[pid].clientId === clientId; }) : [];
    var others = ids.filter(function(pid){ return mine.indexOf(pid) < 0; });
    if (mine.length){
      var g1 = document.createElement("optgroup");
      g1.label = "Projetos de " + clientName(clientId);
      mine.forEach(function(pid){ g1.appendChild(option(pid)); });
      sel.appendChild(g1);
    }
    if (others.length){
      var g2 = document.createElement("optgroup");
      g2.label = mine.length ? "Outros projetos" : "Projetos";
      others.forEach(function(pid){ g2.appendChild(option(pid)); });
      sel.appendChild(g2);
    }
    if (!ids.length) none.textContent = "Sem projetos criados";
    sel.value = selectedId && projects[selectedId] ? selectedId : "";
  }
  // Choosing a project of a client while the task has no client yet fills in that client.
  function onTaskProjectChange(){
    var pid = document.getElementById("taskProject").value;
    var clientSel = document.getElementById("taskClient");
    var p = pid && projects[pid];
    if (p && p.clientId && clients[p.clientId] && (!clientSel.value || clientSel.value === "__new__")){
      clientSel.value = p.clientId;
      document.getElementById("taskClientError").hidden = true;
      document.getElementById("taskNewClientRow").hidden = true;
      populateTaskProjectSelect(pid);
    }
  }

  function onTaskClientChange(){
    var sel = document.getElementById("taskClient");
    var row = document.getElementById("taskNewClientRow");
    document.getElementById("taskClientError").hidden = true;
    if (sel.value !== "__new__") populateTaskProjectSelect(document.getElementById("taskProject").value);
    if (sel.value === "__new__"){
      row.hidden = false;
      document.getElementById("taskNewClientName").focus();
    } else {
      row.hidden = true;
    }
  }
  function createClientFromTaskModal(){
    var input = document.getElementById("taskNewClientName");
    var name = input.value.trim();
    if (!name){ input.focus(); return; }
    var id = addClient(name);
    populateTaskClientSelect(id);
    populateTaskProjectSelect(document.getElementById("taskProject").value);
    input.value = "";
    document.getElementById("taskNewClientRow").hidden = true;
    document.getElementById("newPointInput").focus();
  }

  // ---------- work points (checklist inside a task) ----------
  var pointsDraft = [];
  function renderPointsDraft(){
    var list = document.getElementById("taskPointsList");
    list.innerHTML = "";
    if (pointsDraft.length === 0){ list.appendChild(el("div", "hint", "Sem pontos de trabalho. Escreve abaixo e carrega Enter.")); return; }
    pointsDraft.forEach(function(item){
      var row = el("div", "action-item");
      var cb = document.createElement("input");
      cb.type = "checkbox"; cb.checked = !!item.done;
      cb.setAttribute("aria-label", "Ponto feito");
      cb.addEventListener("change", function(){ item.done = cb.checked; renderPointsDraft(); });
      row.appendChild(cb);
      var text = el("span", "action-text" + (item.done ? " done" : ""));
      text.textContent = item.text;
      row.appendChild(text);
      var rm = el("button", "icon-btn", "×");
      rm.type = "button"; rm.title = "Remover ponto";
      rm.addEventListener("click", function(){
        pointsDraft = pointsDraft.filter(function(a){ return a.id !== item.id; });
        renderPointsDraft();
      });
      row.appendChild(rm);
      list.appendChild(row);
    });
    list.scrollTop = list.scrollHeight;
  }
  function addPointDraft(){
    var input = document.getElementById("newPointInput");
    if (!input.value.trim()) return;
    pointsDraft.push({ id:genId(), text:input.value.trim(), done:false });
    input.value = "";
    renderPointsDraft();
  }
  function closeTaskModal(){ document.getElementById("taskModalOverlay").hidden = true; editingTaskId = null; }

  // ---------- meeting modal ----------
  function openMeetingModal(id){
    editingMeetingId = id || null;
    var overlay = document.getElementById("meetingModalOverlay");
    var title = document.getElementById("meetingModalTitle");
    var deleteBtn = document.getElementById("deleteMeetingBtn");
    if (id && meetings[id]){
      var m = meetings[id];
      title.textContent = "Editar reunião";
      document.getElementById("meetingTitle").value = m.title;
      document.getElementById("meetingDate").value = m.date || todayStr();
      populateClientSelect("meetingClient", m.clientId);
      document.getElementById("meetingAgenda").value = m.agenda || "";
      document.getElementById("meetingNotes").value = m.notes || "";
      document.getElementById("meetingSummary").value = m.summary || "";
      actionsDraft = (m.actions || []).map(function(a){ return { id:a.id, text:a.text, done:!!a.done }; });
      deleteBtn.hidden = false;
      resetConfirm(deleteBtn, "Eliminar");
    } else {
      title.textContent = "Nova reunião";
      document.getElementById("meetingTitle").value = "";
      document.getElementById("meetingDate").value = todayStr();
      populateClientSelect("meetingClient", activeClientFilter);
      document.getElementById("meetingAgenda").value = "";
      document.getElementById("meetingNotes").value = "";
      document.getElementById("meetingSummary").value = "";
      actionsDraft = [];
      deleteBtn.hidden = true;
    }
    document.getElementById("newActionInput").value = "";
    renderActionsDraft();
    overlay.hidden = false;
    setTimeout(function(){ document.getElementById("meetingTitle").focus(); }, 0);
  }
  function closeMeetingModal(){
    document.getElementById("meetingModalOverlay").hidden = true;
    editingMeetingId = null;
    actionsDraft = [];
  }

  // ---------- project modal ----------
  function openProjectModal(id){
    editingProjectId = id || null;
    var overlay = document.getElementById("projectModalOverlay");
    var title = document.getElementById("projectModalTitle");
    var deleteBtn = document.getElementById("deleteProjectBtn");
    var statusEl = document.getElementById("documentAddStatus");
    statusEl.textContent = "";
    statusEl.classList.remove("error");
    documentsPendingCleanup = [];

    if (id && projects[id]){
      var p = projects[id];
      title.textContent = "Editar projeto";
      document.getElementById("projectName").value = p.name;
      populateClientSelect("projectClient", p.clientId);
      setSegValue("projectStatusSeg", p.status || "ativo");
      document.getElementById("projectDescription").value = p.description || "";
      documentsDraft = (p.documents || []).map(function(d){ return Object.assign({}, d); });
      deleteBtn.hidden = false;
      resetConfirm(deleteBtn, "Eliminar");
    } else {
      title.textContent = "Novo projeto";
      document.getElementById("projectName").value = "";
      populateClientSelect("projectClient", activeClientFilter);
      setSegValue("projectStatusSeg", "ativo");
      document.getElementById("projectDescription").value = "";
      documentsDraft = [];
      deleteBtn.hidden = true;
    }
    renderDocumentsDraft();
    overlay.hidden = false;
    setTimeout(function(){ document.getElementById("projectName").focus(); }, 0);
  }
  function closeProjectModal(){
    document.getElementById("projectModalOverlay").hidden = true;
    editingProjectId = null;
    documentsDraft = [];
    if (documentsPendingCleanup.length){
      documentsPendingCleanup.forEach(function(id){ idbDelete(id).catch(function(){}); });
    }
    documentsPendingCleanup = [];
  }

  // ---------- expense modal ----------
  function setReceiptScanStatus(message, isError){
    var status = document.getElementById("receiptScanStatus");
    status.textContent = message;
    status.classList.toggle("error", !!isError);
  }
  function scanReceipt(file){
    var button = document.getElementById("scanReceiptBtn");
    button.disabled = true;
    setReceiptScanStatus("A preparar a leitura do recibo…", false);
    return window.ReceiptScanner.scan(file, function(progress){
      if (!progress || !progress.status) return;
      var labels = {
        "loading tesseract core":"A carregar o motor OCR…",
        "loading language traineddata":"A carregar o modelo português…",
        "initializing tesseract":"A iniciar o motor OCR…",
        "initializing api":"A preparar o reconhecimento…",
        "recognizing text":"A ler o recibo…"
      };
      var label = labels[progress.status];
      if (label){
        var percent = Number.isFinite(progress.progress) ? " " + Math.round(progress.progress * 100) + "%" : "";
        setReceiptScanStatus(label + percent, false);
      }
    }).then(function(suggestion){
      if (suggestion.amount) document.getElementById("expenseAmount").value = suggestion.amount.toFixed(2).replace(".", ",");
      if (suggestion.supplier) document.getElementById("expenseSupplier").value = suggestion.supplier;
      if (suggestion.description) document.getElementById("expenseDescription").value = suggestion.description;
      var found = [];
      if (suggestion.amount) found.push("valor");
      if (suggestion.description) found.push("descrição");
      setReceiptScanStatus(found.length
        ? "Sugestão preenchida (" + found.join(" e ") + "). Revê ou edita antes de guardar."
        : "Não foi possível sugerir campos. Podes preencher manualmente.", false);
    }).catch(function(error){
      console.error("receipt scan failed", error);
      setReceiptScanStatus(error && error.message ? error.message : "Não foi possível ler o recibo. Podes preencher manualmente.", true);
    }).then(function(){
      button.disabled = false;
    });
  }

  function openExpenseModal(id, presetDate){
    editingExpenseId = id || null;
    var overlay = document.getElementById("expenseModalOverlay");
    var title = document.getElementById("expenseModalTitle");
    var deleteBtn = document.getElementById("deleteExpenseBtn");
    if (id && expenses[id]){
      var x = expenses[id];
      title.textContent = "Editar despesa";
      document.getElementById("expenseDate").value = x.date || todayStr();
      document.getElementById("expenseAmount").value = (Number(x.amount)||0).toFixed(2).replace(".", ",");
      setSegValue("expenseCategorySeg", expenseCategory(x));
      populateClientSelect("expenseClient", x.clientId);
      document.getElementById("expenseSupplier").value = x.supplier || "";
      document.getElementById("expenseDescription").value = x.description || "";
      fillFuelFields(x);
      fillMealFields(x);
      deleteBtn.hidden = false;
      resetConfirm(deleteBtn, "Eliminar");
    } else {
      title.textContent = "Nova despesa";
      document.getElementById("expenseDate").value = presetDate || todayStr();
      document.getElementById("expenseAmount").value = "";
      setSegValue("expenseCategorySeg", "outros");
      populateClientSelect("expenseClient", activeClientFilter);
      document.getElementById("expenseSupplier").value = "";
      document.getElementById("expenseDescription").value = "";
      fillFuelFields(null);
      fillMealFields(null);
      deleteBtn.hidden = true;
    }
    setReceiptScanStatus("A foto sugere os campos; confirma ou edita antes de guardar.", false);
    document.getElementById("expenseAmountError").hidden = true;
    overlay.hidden = false;
    setTimeout(function(){ document.getElementById("expenseAmount").focus(); }, 0);
  }
  function closeExpenseModal(){
    document.getElementById("expenseModalOverlay").hidden = true;
    editingExpenseId = null;
  }

  function resetConfirm(btn, label){
    btn.textContent = label;
    btn.classList.remove("confirming");
    btn.dataset.confirming = "";
  }

  // ---------- clients modal ----------
  function openClientsModal(){ renderClientsList(); document.getElementById("clientsModalOverlay").hidden = false; }
  function closeClientsModal(){ document.getElementById("clientsModalOverlay").hidden = true; render(); }

  function renderClientsList(){
    var list = document.getElementById("clientsList");
    list.innerHTML = "";
    var ids = Object.keys(clients).sort(function(a,b){ return (clients[a].order||0) - (clients[b].order||0); });
    if (ids.length === 0) list.appendChild(el("p", "hint", "Ainda não tens clientes. Adiciona um abaixo."));
    ids.forEach(function(cid){
      var row = el("div", "client-row");
      var swatch = el("button", "swatch"); swatch.type = "button"; swatch.style.background = clients[cid].color;
      row.appendChild(swatch);

      var nameInput = el("input");
      nameInput.type = "text"; nameInput.value = clients[cid].name; nameInput.maxLength = 60;
      nameInput.addEventListener("change", function(){ renameClient(cid, nameInput.value); });
      row.appendChild(nameInput);

      var delBtn = el("button", "btn btn-danger btn-sm", "Eliminar");
      delBtn.addEventListener("click", function(){
        if (delBtn.dataset.confirming === "1"){ deleteClient(cid); renderClientsList(); render(); }
        else {
          delBtn.dataset.confirming = "1"; delBtn.textContent = "Confirmar?"; delBtn.classList.add("confirming");
          setTimeout(function(){ delBtn.dataset.confirming = ""; delBtn.textContent = "Eliminar"; delBtn.classList.remove("confirming"); }, 3000);
        }
      });
      row.appendChild(delBtn);

      var picker = el("div", "swatch-picker");
      PALETTE.forEach(function(color){
        var mini = el("button", "swatch-mini" + (color === clients[cid].color ? " selected" : ""));
        mini.type = "button"; mini.style.background = color;
        mini.addEventListener("click", function(){ recolorClient(cid, color); });
        picker.appendChild(mini);
      });

      var wrap = el("div", ""); wrap.style.width = "100%";
      wrap.appendChild(row); wrap.appendChild(picker);
      list.appendChild(wrap);
    });
  }

  // ---------- export / import (backup) ----------
  function exportData(){
    var payload = { clients:clients, tasks:tasks, meetings:meetings, projects:projects, expenses:expenses, exportedAt:new Date().toISOString() };
    var blob = new Blob([JSON.stringify(payload, null, 2)], { type:"application/json" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "quadro-tarefas-copia-" + todayStr() + ".json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function(){ URL.revokeObjectURL(url); }, 1000);
  }

  function importData(file){
    var reader = new FileReader();
    reader.onload = function(){
      var parsed;
      try { parsed = JSON.parse(reader.result); }
      catch(e){ alert("Este ficheiro não é uma cópia de segurança válida."); return; }
      if (!parsed || typeof parsed !== "object" || !parsed.clients || !parsed.tasks){
        alert("Este ficheiro não parece ser uma cópia de segurança deste quadro.");
        return;
      }
      var nClients = Object.keys(parsed.clients||{}).length;
      var nTasks = Object.keys(parsed.tasks||{}).length;
      var nMeetings = Object.keys(parsed.meetings||{}).length;
      var nProjects = Object.keys(parsed.projects||{}).length;
      var nExpenses = Object.keys(parsed.expenses||{}).length;
      var ok = window.confirm(
        "Importar " + nClients + " cliente(s), " + nTasks + " tarefa(s), " + nMeetings + " reunião(ões), " +
        nProjects + " projeto(s) e " + nExpenses + " despesa(s)?\n" +
        "Isto substitui todos os dados atuais deste navegador.\n" +
        "Nota: os documentos anexados aos projetos não fazem parte desta cópia e têm de ser voltados a anexar."
      );
      if (!ok) return;
      clients = parsed.clients || {};
      tasks = parsed.tasks || {};
      meetings = parsed.meetings || {};
      projects = parsed.projects || {};
      expenses = parsed.expenses || {};
      saveStore();
      render();
    };
    reader.readAsText(file);
  }

  // ---------- view switching ----------
  function switchView(view){
    currentView = view;
    Array.prototype.forEach.call(document.querySelectorAll("#viewTabs button"), function(btn){
      btn.classList.toggle("active", btn.dataset.view === view);
      if (btn.dataset.view === view) btn.setAttribute("aria-current", "page");
      else btn.removeAttribute("aria-current");
    });
    var labels = { tasks:"+ Nova tarefa", meetings:"+ Nova reunião", projects:"+ Novo projeto", expenses:"+ Nova despesa" };
    document.getElementById("newItemBtn").textContent = labels[view] || "+ Novo";
    render();
    if (view === "home") loadFolderData(); // pick up changes made to the Excel files
  }

  // ---------- wiring ----------
  function wireStaticUI(){
    Array.prototype.forEach.call(document.querySelectorAll("#viewTabs button"), function(btn){
      btn.addEventListener("click", function(){ switchView(btn.dataset.view); });
    });

    document.getElementById("newItemBtn").addEventListener("click", function(){
      if (currentView === "tasks") openTaskModal(null);
      else if (currentView === "meetings") openMeetingModal(null);
      else if (currentView === "projects") openProjectModal(null);
      else openExpenseModal(null);
    });
    document.getElementById("taskQuickAddBtn").addEventListener("click", openNewTaskFromList);
    var titleInput = document.getElementById("taskTitle");
    titleInput.addEventListener("input", renderQuickPreview);
    titleInput.addEventListener("blur", applyTitleShortcuts);
    titleInput.addEventListener("keydown", function(e){
      // Enter in the title of a new task: apply the shortcuts and move on to the client.
      if (e.key === "Enter" && !editingTaskId){
        e.preventDefault();
        applyTitleShortcuts();
        document.getElementById(document.getElementById("taskClient").value ? "newPointInput" : "taskClient").focus();
      }
    });
    Array.prototype.forEach.call(document.querySelectorAll("#taskGroupingSeg button"), function(b){
      b.addEventListener("click", function(){
        taskGrouping = b.dataset.val;
        saveTaskViewPrefs();
        renderBoard();
      });
    });
    document.getElementById("taskShowDone").addEventListener("change", function(){
      showDoneByClient = this.checked;
      saveTaskViewPrefs();
      renderBoard();
    });

    document.getElementById("taskClient").addEventListener("change", onTaskClientChange);
    document.getElementById("taskProject").addEventListener("change", onTaskProjectChange);
    document.getElementById("taskNewClientAddBtn").addEventListener("click", createClientFromTaskModal);
    document.getElementById("taskNewClientName").addEventListener("keydown", function(e){
      if (e.key === "Enter"){ e.preventDefault(); createClientFromTaskModal(); }
    });
    document.getElementById("addPointBtn").addEventListener("click", function(){
      addPointDraft();
      document.getElementById("newPointInput").focus();
    });
    document.getElementById("newPointInput").addEventListener("keydown", function(e){
      if (e.key === "Enter"){ e.preventDefault(); addPointDraft(); }
    });
    document.getElementById("homePrevMonthBtn").addEventListener("click", function(){ shiftHomeMonth(-1); });
    document.getElementById("homeNextMonthBtn").addEventListener("click", function(){ shiftHomeMonth(1); });
    document.getElementById("addMeetingBottomBtn").addEventListener("click", function(){ openMeetingModal(null); });
    document.getElementById("addProjectBottomBtn").addEventListener("click", function(){ openProjectModal(null); });
    document.getElementById("addExpenseBottomBtn").addEventListener("click", function(){ openExpenseModal(null); });
    document.getElementById("exportExpensesBtn").addEventListener("click", exportExpensesCSV);
    document.getElementById("prevMonthBtn").addEventListener("click", function(){ shiftMonth(-1); });
    document.getElementById("nextMonthBtn").addEventListener("click", function(){ shiftMonth(1); });
    document.getElementById("closeDayExpensesBtn").addEventListener("click", closeDayExpensesModal);
    document.getElementById("addDayExpenseBtn").addEventListener("click", function(){
      var d = dayModalDate;
      closeDayExpensesModal();
      openExpenseModal(null, d);
    });
    document.getElementById("dayExpensesModalOverlay").addEventListener("click", function(e){
      if (e.target === this) closeDayExpensesModal();
    });

    document.getElementById("cancelTaskBtn").addEventListener("click", closeTaskModal);
    document.getElementById("saveTaskBtn").addEventListener("click", saveTaskFromModal);
    document.getElementById("deleteTaskBtn").addEventListener("click", function(){
      var btn = document.getElementById("deleteTaskBtn");
      if (btn.dataset.confirming === "1"){ deleteEditingTask(); }
      else { btn.dataset.confirming = "1"; btn.textContent = "Confirmar eliminação?"; btn.classList.add("confirming"); }
    });
    document.getElementById("taskModalOverlay").addEventListener("click", function(e){ if (e.target === this) closeTaskModal(); });

    document.getElementById("cancelMeetingBtn").addEventListener("click", closeMeetingModal);
    document.getElementById("saveMeetingBtn").addEventListener("click", saveMeetingFromModal);
    document.getElementById("deleteMeetingBtn").addEventListener("click", function(){
      var btn = document.getElementById("deleteMeetingBtn");
      if (btn.dataset.confirming === "1"){ deleteEditingMeeting(); }
      else { btn.dataset.confirming = "1"; btn.textContent = "Confirmar eliminação?"; btn.classList.add("confirming"); }
    });
    document.getElementById("meetingModalOverlay").addEventListener("click", function(e){ if (e.target === this) closeMeetingModal(); });
    document.getElementById("addActionBtn").addEventListener("click", function(){
      var input = document.getElementById("newActionInput");
      addActionDraft(input.value); input.value = ""; input.focus();
    });
    document.getElementById("newActionInput").addEventListener("keydown", function(e){
      if (e.key === "Enter" && this.value.trim()){ addActionDraft(this.value); this.value = ""; }
    });

    document.getElementById("cancelProjectBtn").addEventListener("click", closeProjectModal);
    document.getElementById("saveProjectBtn").addEventListener("click", saveProjectFromModal);
    document.getElementById("deleteProjectBtn").addEventListener("click", function(){
      var btn = document.getElementById("deleteProjectBtn");
      if (btn.dataset.confirming === "1"){ deleteEditingProject(); }
      else { btn.dataset.confirming = "1"; btn.textContent = "Confirmar eliminação?"; btn.classList.add("confirming"); }
    });
    document.getElementById("projectModalOverlay").addEventListener("click", function(e){ if (e.target === this) closeProjectModal(); });
    document.getElementById("addDocumentBtn").addEventListener("click", function(){ document.getElementById("documentFileInput").click(); });
    document.getElementById("documentFileInput").addEventListener("change", function(e){
      addDocumentsFromFiles(e.target.files);
      e.target.value = "";
    });

    document.getElementById("cancelExpenseBtn").addEventListener("click", closeExpenseModal);
    document.getElementById("saveExpenseBtn").addEventListener("click", saveExpenseFromModal);
    document.getElementById("scanReceiptBtn").addEventListener("click", function(){ document.getElementById("receiptFileInput").click(); });
    document.getElementById("receiptFileInput").addEventListener("change", function(e){
      var file = e.target.files && e.target.files[0];
      if (file) scanReceipt(file);
      e.target.value = "";
    });
    document.getElementById("deleteExpenseBtn").addEventListener("click", function(){
      var btn = document.getElementById("deleteExpenseBtn");
      if (btn.dataset.confirming === "1"){ deleteEditingExpense(); }
      else { btn.dataset.confirming = "1"; btn.textContent = "Confirmar eliminação?"; btn.classList.add("confirming"); }
    });
    document.getElementById("expenseModalOverlay").addEventListener("click", function(e){ if (e.target === this) closeExpenseModal(); });

    ["prioritySeg","statusSeg","projectStatusSeg","expenseCategorySeg","expensePaymentSeg","mealWithSeg"].forEach(function(segId){
      document.getElementById(segId).addEventListener("click", function(e){
        var btn = e.target.closest("button");
        if (!btn) return;
        setSegValue(segId, btn.dataset.val);
        if (segId === "expenseCategorySeg" || segId === "expensePaymentSeg" || segId === "mealWithSeg"){
          updateFuelFields();
          updateMealFields();
        }
      });
    });
    document.getElementById("excelSyncRetryBtn").addEventListener("click", function(){ runExcelSync(true); });
    document.getElementById("expenseAmount").addEventListener("input", function(){
      document.getElementById("expenseAmountError").hidden = true;
    });
    document.getElementById("expenseClient").addEventListener("change", function(){
      document.getElementById("expenseClientError").hidden = true;
    });
    document.getElementById("mealSupplier").addEventListener("input", function(){
      document.getElementById("mealSupplierError").hidden = true;
    });

    document.getElementById("manageClientsBtn").addEventListener("click", openClientsModal);
    document.getElementById("closeClientsBtn").addEventListener("click", closeClientsModal);
    document.getElementById("clientsModalOverlay").addEventListener("click", function(e){ if (e.target === this) closeClientsModal(); });
    document.getElementById("addClientBtn").addEventListener("click", function(){
      var input = document.getElementById("newClientName");
      if (input.value.trim()){ addClient(input.value.trim()); input.value = ""; }
    });
    document.getElementById("newClientName").addEventListener("keydown", function(e){
      if (e.key === "Enter" && this.value.trim()){ addClient(this.value.trim()); this.value = ""; }
    });

    document.getElementById("exportBtn").addEventListener("click", exportData);
    document.getElementById("importBtn").addEventListener("click", function(){ document.getElementById("importFileInput").click(); });
    document.getElementById("importFileInput").addEventListener("change", function(e){
      var file = e.target.files && e.target.files[0];
      if (file) importData(file);
      e.target.value = "";
    });

    var dismissedWarn = false;
    try { dismissedWarn = localStorage.getItem("qt_storage_warning_dismissed"); } catch(e){}
    if (!dismissedWarn) document.getElementById("storageWarning").hidden = false;
    document.getElementById("dismissWarning").addEventListener("click", function(){
      document.getElementById("storageWarning").hidden = true;
      try { localStorage.setItem("qt_storage_warning_dismissed", "1"); } catch(e){}
    });

    try {
      var isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
      var isStandalone = window.navigator.standalone === true;
      var installDismissed = false;
      try { installDismissed = localStorage.getItem("qt_install_banner_dismissed"); } catch(e){}
      var installBanner = document.getElementById("installAppBanner");
      if (isIOS && !isStandalone && !installDismissed && installBanner){
        installBanner.hidden = false;
      }
      var dismissInstallBtn = document.getElementById("dismissInstallBanner");
      if (dismissInstallBtn){
        dismissInstallBtn.addEventListener("click", function(){
          installBanner.hidden = true;
          try { localStorage.setItem("qt_install_banner_dismissed", "1"); } catch(e){}
        });
      }
    } catch(e){ /* deteção de iOS falhou, não é crítico */ }

    document.addEventListener("keydown", function(e){
      if (e.key === "Escape"){
        if (!document.getElementById("taskModalOverlay").hidden) closeTaskModal();
        if (!document.getElementById("meetingModalOverlay").hidden) closeMeetingModal();
        if (!document.getElementById("projectModalOverlay").hidden) closeProjectModal();
        if (!document.getElementById("expenseModalOverlay").hidden) closeExpenseModal();
        if (!document.getElementById("dayExpensesModalOverlay").hidden) closeDayExpensesModal();
        if (!document.getElementById("clientsModalOverlay").hidden) closeClientsModal();
      }
    });
  }

  function showFatalError(where, err){
    try {
      var banner = document.getElementById("fatalErrorBanner");
      if (!banner) return;
      banner.hidden = false;
      var span = banner.querySelector("span");
      if (span) span.textContent = "Erro em " + where + ": " + (err && err.message ? err.message : err) + ". Tenta abrir o ficheiro pelo Safari (Partilhar > Abrir no Safari) em vez de o pré-visualizar.";
      var btn = document.getElementById("dismissFatalError");
      if (btn && !btn._wired){ btn._wired = true; btn.addEventListener("click", function(){ banner.hidden = true; }); }
    } catch(e2){ /* nada mais a fazer */ }
  }

  function main(){
    try { loadStore(); }
    catch(e){ console.error("loadStore falhou", e); showFatalError("carregar dados", e); }
    forgetStoredSalesCopy();

    try { wireStaticUI(); }
    catch(e){ console.error("wireStaticUI falhou", e); showFatalError("ligar botões", e); }

    try { render(); }
    catch(e){ console.error("render falhou", e); showFatalError("mostrar conteúdo", e); }

    try { loadFolderData(); }
    catch(e){ console.error("loadFolderData falhou", e); }
    scheduleExcelSync();

    registerServiceWorker();
  }

  // Offline support when served over http(s); skipped for file:// opens.
  function registerServiceWorker(){
    if (!("serviceWorker" in navigator) || !/^https?:$/.test(location.protocol)) return;
    navigator.serviceWorker.register("sw.js").catch(function(err){
      console.error("service worker registration failed", err);
    });
  }

  // Started by auth.js once the user is signed in (or immediately in local-only mode).
  var started = false;
  // opts.userId: signed-in user's id (null/undefined = local-only mode, legacy key).
  window.QuadroApp = {
    getSyncData: function(){
      return JSON.parse(JSON.stringify({ clients:clients, tasks:tasks, meetings:meetings, projects:projects, expenses:expenses }));
    },
    applySyncData: function(data, baseline){
      if (!data || typeof data !== "object") return;
      var latest = this.getSyncData();
      var concurrentChanges = false;
      applyingSyncedData = true;
      clients = data.clients || {};
      tasks = data.tasks || {};
      meetings = data.meetings || {};
      projects = data.projects || {};
      expenses = data.expenses || {};
      if (baseline){
        var target = { clients:clients, tasks:tasks, meetings:meetings, projects:projects, expenses:expenses };
        ["clients","tasks","meetings","projects","expenses"].forEach(function(collection){
          var ids = {};
          Object.keys(baseline[collection] || {}).forEach(function(id){ ids[id] = true; });
          Object.keys(latest[collection] || {}).forEach(function(id){ ids[id] = true; });
          Object.keys(ids).forEach(function(id){
            var before = baseline[collection] && baseline[collection][id];
            var after = latest[collection] && latest[collection][id];
            if (JSON.stringify(before) === JSON.stringify(after)) return;
            concurrentChanges = true;
            if (after) target[collection][id] = after;
            else delete target[collection][id];
          });
        });
      }
      saveStore();
      applyingSyncedData = false;
      render();
      return concurrentChanges;
    },
    start: function(opts){
      if (started) return;
      started = true;
      currentUserId = (opts && opts.userId) || null;
      if (currentUserId) STORAGE_KEY = LEGACY_STORAGE_KEY + ":" + currentUserId;
      main();
    }
  };
})();
