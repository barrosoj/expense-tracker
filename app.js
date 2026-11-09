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
  var EXPENSE_CATEGORIES = ["deslocacoes","refeicoes","alojamento","material"];
  var EXPENSE_CATEGORY_LABEL = { deslocacoes:"Deslocações", refeicoes:"Refeições", alojamento:"Alojamento", material:"Material" };
  var DASHBOARD_CLIENTS = ["Kennametal","Bass","Arno","GAIt"];

  var currentView = "home"; // "home" | "tasks" | "meetings" | "projects" | "expenses" | "clients"

  var clients = {};
  var tasks = {};
  var meetings = {};
  var projects = {};
  var expenses = {};
  var monthlyValues = {};
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
  var calendarSelectedDate = todayStr();
  var receiptPhotoDraft = null;
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
  function formatCurrency(n){
    var v = Number(n) || 0;
    try { return v.toLocaleString("pt-PT", { style:"currency", currency:"EUR" }); }
    catch(e){ return v.toFixed(2) + " €"; }
  }
  function parseAmountInput(str){
    if (!str) return 0;
    var cleaned = String(str).replace(/[€\s]/g, "").replace(/\./g, "").replace(",", ".");
    var n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  }

  // ---------- IndexedDB (document and receipt photo blobs) ----------
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
      expenses: {},
      monthlyValues: {}
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
        monthlyValues = parsed.monthlyValues || {};
        return;
      } catch(e){ /* fall through to seed */ }
    }
    var seed = defaultSeed();
    clients = seed.clients;
    tasks = seed.tasks;
    meetings = seed.meetings;
    projects = seed.projects;
    expenses = seed.expenses;
    monthlyValues = seed.monthlyValues;
    saveStore();
  }

  function saveStore(){
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ clients:clients, tasks:tasks, meetings:meetings, projects:projects, expenses:expenses, monthlyValues:monthlyValues }));
      if (!applyingSyncedData && window.QuadroSync && currentUserId){
        window.QuadroSync.localChanged();
      }
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

  function render(){
    renderFilters();
    document.getElementById("dashboardView").hidden = currentView !== "home";
    document.getElementById("board").hidden = currentView !== "tasks";
    document.getElementById("meetingsView").hidden = currentView !== "meetings";
    document.getElementById("projectsView").hidden = currentView !== "projects";
    document.getElementById("expensesView").hidden = currentView !== "expenses";
    document.getElementById("clientsView").hidden = currentView !== "clients";
    if (currentView === "home") renderDashboard();
    else if (currentView === "tasks") renderBoard();
    else if (currentView === "meetings") renderMeetingsList();
    else if (currentView === "projects") renderProjectsList();
    else if (currentView === "expenses") renderExpensesList();
    else renderClientsView();
  }

  function renderFilters(){
    var wrap = document.getElementById("filters");
    wrap.innerHTML = "";
    wrap.hidden = currentView === "home" || currentView === "clients";
    if (wrap.hidden) return;
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

  function dashboardValueKey(month, clientName){
    return month + "|" + clientName.toLowerCase();
  }
  function renderDashboard(){
    var month = todayStr().slice(0,7);
    document.getElementById("dashboardMonthLabel").textContent = monthLabel(month);
    var values = document.getElementById("monthlyValues");
    values.innerHTML = "";
    DASHBOARD_CLIENTS.forEach(function(name){
      var key = dashboardValueKey(month, name);
      var saved = monthlyValues[key] || { sales:0, purchases:0 };
      var row = el("div", "monthly-value-row");
      row.appendChild(el("div", "monthly-value-client", name));
      [["sales","Vendas"],["purchases","Compras"]].forEach(function(field){
        var cell = el("label", "monthly-value-field");
        cell.appendChild(el("span", "", field[1]));
        var input = el("input", "mono");
        input.type = "text";
        input.inputMode = "decimal";
        input.setAttribute("aria-label", field[1] + " — " + name);
        input.value = (Number(saved[field[0]]) || 0).toFixed(2).replace(".", ",");
        input.addEventListener("change", function(){
          var current = monthlyValues[key] || { sales:0, purchases:0 };
          current[field[0]] = Math.max(0, parseAmountInput(input.value));
          current.updatedAt = Date.now();
          monthlyValues[key] = current;
          input.value = current[field[0]].toFixed(2).replace(".", ",");
          saveStore();
        });
        cell.appendChild(input);
        row.appendChild(cell);
      });
      values.appendChild(row);
    });
    renderCalendar(month);
  }

  function calendarItemsForDate(date){
    var items = [];
    Object.keys(tasks).forEach(function(id){
      var task = tasks[id];
      if (task.due === date && task.status !== "done") items.push({ type:"task", id:id, title:task.title });
    });
    Object.keys(meetings).forEach(function(id){
      var meeting = meetings[id];
      if (meeting.date === date) items.push({ type:"meeting", id:id, title:meeting.title });
    });
    return items;
  }

  function renderCalendar(month){
    var parts = month.split("-").map(Number);
    var year = parts[0], monthIndex = parts[1] - 1;
    document.getElementById("calendarMonthLabel").textContent = monthLabel(month);
    var firstDay = new Date(year, monthIndex, 1);
    var startOffset = (firstDay.getDay() + 6) % 7;
    var daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
    var selected = calendarSelectedDate.slice(0,7) === month ? calendarSelectedDate : month + "-01";
    calendarSelectedDate = selected;
    var grid = document.getElementById("calendarGrid");
    grid.innerHTML = "";
    ["Seg","Ter","Qua","Qui","Sex","Sáb","Dom"].forEach(function(label){
      grid.appendChild(el("div", "calendar-weekday", label));
    });
    for (var blank = 0; blank < startOffset; blank++) grid.appendChild(el("div", "calendar-day calendar-day-empty"));
    for (var day = 1; day <= daysInMonth; day++){
      var date = year + "-" + String(monthIndex+1).padStart(2,"0") + "-" + String(day).padStart(2,"0");
      var items = calendarItemsForDate(date);
      var button = el("button", "calendar-day" + (date === calendarSelectedDate ? " selected" : "") + (date === todayStr() ? " today" : ""));
      button.type = "button";
      button.setAttribute("aria-label", formatDateShort(date) + (items.length ? ", " + items.length + " eventos" : ""));
      button.appendChild(el("span", "calendar-day-number", String(day)));
      if (items.length) button.appendChild(el("span", "calendar-day-count mono", String(items.length)));
      button.addEventListener("click", function(selectedDate){
        return function(){ calendarSelectedDate = selectedDate; renderCalendar(month); };
      }(date));
      grid.appendChild(button);
    }

    document.getElementById("calendarSelectedLabel").textContent = formatDateShort(selected);
    var eventList = document.getElementById("calendarEvents");
    eventList.innerHTML = "";
    var selectedItems = calendarItemsForDate(selected);
    if (!selectedItems.length){
      eventList.appendChild(el("div", "empty-hint", "Sem tarefas ou reuniões neste dia."));
      return;
    }
    selectedItems.forEach(function(item){
      var button = el("button", "calendar-event");
      button.type = "button";
      button.appendChild(el("span", "calendar-event-type", item.type === "task" ? "Tarefa" : "Reunião"));
      button.appendChild(el("span", "calendar-event-title", escapeHtml(item.title || "")));
      button.addEventListener("click", function(){
        if (item.type === "task") openTaskModal(item.id);
        else openMeetingModal(item.id);
      });
      eventList.appendChild(button);
    });
  }

  function renderClientsView(){
    var wrap = document.getElementById("clientWorkspaces");
    wrap.innerHTML = "";
    var ids = Object.keys(clients).sort(function(a,b){ return (clients[a].order||0) - (clients[b].order||0); });
    if (!ids.length){
      wrap.appendChild(el("div", "empty-hint", "Ainda não tens clientes. Usa «Gerir clientes» para adicionar o primeiro."));
      return;
    }
    ids.forEach(function(clientId){
      var workspace = el("section", "client-workspace");
      var header = el("div", "client-workspace-header");
      var title = el("h2", "client-workspace-title");
      var dot = el("span", "client-dot");
      dot.style.background = clients[clientId].color;
      title.appendChild(dot);
      title.appendChild(document.createTextNode(clients[clientId].name));
      header.appendChild(title);
      [["Tarefa",openTaskModal],["Reunião",openMeetingModal],["Projeto",openProjectModal]].forEach(function(action){
        var button = el("button", "btn btn-ghost btn-sm", "+ " + action[0]);
        button.type = "button";
        button.addEventListener("click", function(){ activeClientFilter = clientId; action[1](null); });
        header.appendChild(button);
      });
      workspace.appendChild(header);
      [["Tarefas",tasks,"title","due",openTaskModal],["Reuniões",meetings,"title","date",openMeetingModal],["Projetos",projects,"name","",openProjectModal]].forEach(function(group){
        var section = el("div", "client-workspace-group");
        section.appendChild(el("h3", "", group[0]));
        var list = el("div", "client-workspace-items");
        var records = Object.keys(group[1]).filter(function(id){ return group[1][id].clientId === clientId; });
        records.sort(function(a,b){
          var da = group[1][a][group[3]] || "", db = group[1][b][group[3]] || "";
          return da.localeCompare(db);
        });
        if (!records.length) list.appendChild(el("div", "hint", "Sem registos."));
        records.forEach(function(id){
          var record = group[1][id];
          var item = el("button", "client-workspace-item");
          item.type = "button";
          item.appendChild(el("span", "client-workspace-item-title", escapeHtml(record[group[2]] || "")));
          if (group[3] && record[group[3]]) item.appendChild(el("span", "hint", formatDateShort(record[group[3]])));
          item.addEventListener("click", function(){ group[4](id); });
          list.appendChild(item);
        });
        section.appendChild(list);
        workspace.appendChild(section);
      });
      wrap.appendChild(workspace);
    });
  }

  // ---------- tasks board ----------
  function sortedIdsForStatus(statusId, ids){
    return ids.filter(function(id){ return tasks[id].status === statusId; }).sort(function(a,b){
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
    var board = document.getElementById("board");
    board.innerHTML = "";
    var visIds = visibleIds();

    STATUSES.forEach(function(status){
      var col = el("div", "column");
      var head = el("div", "column-head");
      var left = el("div", "column-head-left");
      left.appendChild(el("span", "column-title", status.label));
      var ids = sortedIdsForStatus(status.id, visIds);
      left.appendChild(el("span", "column-count mono", String(ids.length)));
      head.appendChild(left);
      col.appendChild(head);

      var quickAdd = el("div", "quick-add");
      var input = el("input");
      input.type = "text";
      input.placeholder = "Adicionar tarefa rápida...";
      input.maxLength = 140;
      input.addEventListener("keydown", function(e){
        if (e.key === "Enter" && input.value.trim()){
          quickAddTask(input.value.trim(), status.id);
          input.value = "";
        }
      });
      quickAdd.appendChild(input);
      col.appendChild(quickAdd);

      var zone = el("div", "column-dropzone");
      zone.addEventListener("dragover", function(e){ e.preventDefault(); zone.classList.add("drag-over"); });
      zone.addEventListener("dragleave", function(){ zone.classList.remove("drag-over"); });
      zone.addEventListener("drop", function(e){
        e.preventDefault();
        zone.classList.remove("drag-over");
        if (dragTaskId) setTaskStatus(dragTaskId, status.id);
      });

      if (ids.length === 0){
        zone.appendChild(el("div", "empty-hint", "Sem tarefas aqui."));
      } else {
        ids.forEach(function(id){ zone.appendChild(renderCard(id, status.id)); });
      }
      col.appendChild(zone);
      board.appendChild(col);
    });
  }

  function renderCard(id, statusId){
    var t = tasks[id];
    var card = el("div", "card");
    card.draggable = true;
    var color = clientColor(t.clientId);
    if (color) card.style.setProperty("--client-color", color);

    var top = el("div", "card-top");
    if (color){ var dot = el("span", "client-dot"); dot.style.background = color; top.appendChild(dot); }
    var cname = clientName(t.clientId);
    top.appendChild(el("span", "client-name" + (cname ? "" : " none"), escapeHtml(cname || "Sem cliente")));
    top.appendChild(el("span", "priority-badge priority-" + t.priority, PRIORITY_LABEL[t.priority] || "Média"));
    card.appendChild(top);

    card.appendChild(el("div", "card-title", escapeHtml(t.title)));
    if (t.notes) card.appendChild(el("div", "card-notes", escapeHtml(t.notes)));

    var meta = el("div", "card-meta");
    var dueInfo = formatDue(t.due);
    meta.appendChild(el("span", "due mono" + (dueInfo ? " " + dueInfo.cls : ""), dueInfo ? dueInfo.text : ""));

    var move = el("div", "card-move");
    var idx = STATUSES.findIndex(function(s){ return s.id === statusId; });
    var back = el("button", "icon-btn", "‹");
    back.title = "Mover para trás";
    back.disabled = idx === 0;
    back.addEventListener("click", function(e){ e.stopPropagation(); if (idx > 0) setTaskStatus(id, STATUSES[idx-1].id); });
    var fwd = el("button", "icon-btn", "›");
    fwd.title = "Mover para a frente";
    fwd.disabled = idx === STATUSES.length - 1;
    fwd.addEventListener("click", function(e){ e.stopPropagation(); if (idx < STATUSES.length - 1) setTaskStatus(id, STATUSES[idx+1].id); });
    move.appendChild(back); move.appendChild(fwd);
    meta.appendChild(move);
    card.appendChild(meta);

    card.addEventListener("click", function(){ openTaskModal(id); });
    card.addEventListener("dragstart", function(){ dragTaskId = id; card.classList.add("dragging"); });
    card.addEventListener("dragend", function(){ dragTaskId = null; card.classList.remove("dragging"); });
    return card;
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
      var cat = EXPENSE_CATEGORIES.indexOf(x.category) >= 0 ? x.category : "material";
      var amt = Number(x.amount) || 0;
      if (!byDay[x.date]) byDay[x.date] = { deslocacoes:0, refeicoes:0, alojamento:0, material:0, total:0 };
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

    var grand = { deslocacoes:0, refeicoes:0, alojamento:0, material:0, total:0 };
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
      row.appendChild(el("div", "expense-cat-badge expense-cat-" + (x.category || "material"), EXPENSE_CATEGORY_LABEL[x.category] || "Material"));
      var descWrap = el("div", "day-expense-desc");
      descWrap.textContent = x.description || "";
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
    var lines = ["Data;Categoria;Cliente;Fornecedor;Descrição;Valor (€)"];
    ids.forEach(function(id){
      var x = expenses[id];
      lines.push([
        csvField(x.date || ""),
        csvField(EXPENSE_CATEGORY_LABEL[x.category] || ""),
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
  function quickAddTask(title, statusId){
    var id = genId();
    tasks[id] = { title:title, clientId:activeClientFilter||null, priority:"media", status:statusId, due:"", notes:"", createdAt:Date.now(), updatedAt:Date.now() };
    saveStore(); render();
  }
  function setTaskStatus(id, statusId){
    if (!tasks[id] || tasks[id].status === statusId) return;
    tasks[id].status = statusId; tasks[id].updatedAt = Date.now();
    saveStore(); render();
  }
  function saveTaskFromModal(){
    var title = document.getElementById("taskTitle").value.trim();
    if (!title){ document.getElementById("taskTitle").focus(); return; }
    var payload = {
      title:title,
      clientId: document.getElementById("taskClient").value || null,
      priority: document.getElementById("prioritySeg").dataset.value,
      status: document.getElementById("statusSeg").dataset.value,
      due: document.getElementById("taskDue").value || "",
      notes: document.getElementById("taskNotes").value.trim(),
      updatedAt: Date.now()
    };
    if (editingTaskId){ tasks[editingTaskId] = Object.assign({}, tasks[editingTaskId], payload); }
    else { payload.createdAt = Date.now(); tasks[genId()] = payload; }
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
    if (amount <= 0){ document.getElementById("expenseAmount").focus(); return; }
    var date = document.getElementById("expenseDate").value || todayStr();
    var id = editingExpenseId || genId();
    var oldPhotoId = editingExpenseId && expenses[editingExpenseId] ? expenses[editingExpenseId].receiptPhotoId : null;
    var photoId = receiptPhotoDraft ? (oldPhotoId || "receipt_" + id) : oldPhotoId;
    var payload = {
      date: date,
      category: document.getElementById("expenseCategorySeg").dataset.value,
      clientId: document.getElementById("expenseClient").value || null,
      supplier: document.getElementById("expenseSupplier").value.trim(),
      description: description,
      amount: amount,
      receiptPhotoId: photoId || null,
      updatedAt: Date.now()
    };
    var persistPhoto = Promise.resolve();
    if (receiptPhotoDraft){
      persistPhoto = idbPut({
        id:photoId,
        name:receiptPhotoDraft.name,
        contentType:receiptPhotoDraft.type,
        sizeBytes:receiptPhotoDraft.size,
        folder:"receipts/" + date.slice(0,7),
        blob:receiptPhotoDraft
      });
    } else if (photoId && editingExpenseId && expenses[editingExpenseId].date !== date){
      persistPhoto = idbGet(photoId).then(function(photo){
        if (!photo) return;
        photo.folder = "receipts/" + date.slice(0,7);
        return idbPut(photo);
      });
    }
    persistPhoto.then(function(){
      if (editingExpenseId){ expenses[editingExpenseId] = Object.assign({}, expenses[editingExpenseId], payload); }
      else { payload.createdAt = Date.now(); expenses[id] = payload; }
      receiptPhotoDraft = null;
      saveStore(); closeExpenseModal(); render();
    }).catch(function(err){
      console.error("save receipt photo failed", err);
      setReceiptScanStatus("Não foi possível guardar a foto neste dispositivo. Liberta espaço ou tenta outra imagem.", true);
    });
  }
  function deleteEditingExpense(){
    if (!editingExpenseId) return;
    var photoId = expenses[editingExpenseId].receiptPhotoId;
    delete expenses[editingExpenseId];
    if (photoId) idbDelete(photoId).catch(function(err){ console.error("delete receipt photo failed", err); });
    saveStore(); closeExpenseModal(); render();
  }

  // ---------- client ops ----------
  function addClient(name){
    var id = genId();
    clients[id] = { name:name.trim(), color:nextPaletteColor(), order:Date.now(), updatedAt:Date.now() };
    saveStore(); renderClientsList(); render();
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
  function openTaskModal(id, presetStatus){
    editingTaskId = id || null;
    var overlay = document.getElementById("taskModalOverlay");
    var title = document.getElementById("taskModalTitle");
    var deleteBtn = document.getElementById("deleteTaskBtn");
    if (id && tasks[id]){
      var t = tasks[id];
      title.textContent = "Editar tarefa";
      document.getElementById("taskTitle").value = t.title;
      populateClientSelect("taskClient", t.clientId);
      setSegValue("prioritySeg", t.priority || "media");
      setSegValue("statusSeg", t.status || "todo");
      document.getElementById("taskDue").value = t.due || "";
      document.getElementById("taskNotes").value = t.notes || "";
      deleteBtn.hidden = false;
      resetConfirm(deleteBtn, "Eliminar");
    } else {
      title.textContent = "Nova tarefa";
      document.getElementById("taskTitle").value = "";
      populateClientSelect("taskClient", activeClientFilter);
      setSegValue("prioritySeg", "media");
      setSegValue("statusSeg", presetStatus || "todo");
      document.getElementById("taskDue").value = "";
      document.getElementById("taskNotes").value = "";
      deleteBtn.hidden = true;
    }
    overlay.hidden = false;
    setTimeout(function(){ document.getElementById("taskTitle").focus(); }, 0);
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
  function renderReceiptPhoto(expense){
    var info = document.getElementById("receiptPhotoInfo");
    info.innerHTML = "";
    if (receiptPhotoDraft){
      info.textContent = receiptPhotoDraft.name + " — será guardada em receipts/" + (document.getElementById("expenseDate").value || todayStr()).slice(0,7);
      return;
    }
    var photoId = expense && expense.receiptPhotoId;
    if (!photoId) return;
    idbGet(photoId).then(function(photo){
      if (document.getElementById("expenseModalOverlay").hidden) return;
      var currentExpense = editingExpenseId && expenses[editingExpenseId];
      if (!currentExpense || currentExpense.receiptPhotoId !== photoId) return;
      if (!photo){ info.textContent = "Foto não disponível neste dispositivo."; return; }
      var button = el("button", "document-name", escapeHtml(photo.name || "Ver foto do recibo"));
      button.type = "button";
      button.addEventListener("click", function(){ openReceiptPhoto(photoId, photo.name); });
      info.appendChild(button);
      info.appendChild(el("span", "document-meta mono", photo.folder || ""));
    }).catch(function(){ info.textContent = "Foto não disponível neste dispositivo."; });
  }
  function openReceiptPhoto(id, name){
    idbGet(id).then(function(photo){
      if (!photo || !photo.blob){ alert("Não foi possível encontrar esta foto neste dispositivo."); return; }
      var url = URL.createObjectURL(photo.blob);
      window.open(url, "_blank");
      setTimeout(function(){ URL.revokeObjectURL(url); }, 60000);
    }).catch(function(err){
      console.error("openReceiptPhoto failed", err);
      alert("Não foi possível abrir \"" + name + "\".");
    });
  }
  function scanReceipt(file){
    receiptPhotoDraft = file;
    renderReceiptPhoto(null);
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
    receiptPhotoDraft = null;
    var overlay = document.getElementById("expenseModalOverlay");
    var title = document.getElementById("expenseModalTitle");
    var deleteBtn = document.getElementById("deleteExpenseBtn");
    if (id && expenses[id]){
      var x = expenses[id];
      title.textContent = "Editar despesa";
      document.getElementById("expenseDate").value = x.date || todayStr();
      document.getElementById("expenseAmount").value = (Number(x.amount)||0).toFixed(2).replace(".", ",");
      setSegValue("expenseCategorySeg", x.category || "deslocacoes");
      populateClientSelect("expenseClient", x.clientId);
      document.getElementById("expenseSupplier").value = x.supplier || "";
      document.getElementById("expenseDescription").value = x.description || "";
      deleteBtn.hidden = false;
      resetConfirm(deleteBtn, "Eliminar");
    } else {
      title.textContent = "Nova despesa";
      document.getElementById("expenseDate").value = presetDate || todayStr();
      document.getElementById("expenseAmount").value = "";
      setSegValue("expenseCategorySeg", "deslocacoes");
      populateClientSelect("expenseClient", activeClientFilter);
      document.getElementById("expenseSupplier").value = "";
      document.getElementById("expenseDescription").value = "";
      deleteBtn.hidden = true;
    }
    renderReceiptPhoto(id && expenses[id] ? expenses[id] : null);
    setReceiptScanStatus("A foto sugere os campos; confirma ou edita antes de guardar.", false);
    overlay.hidden = false;
    setTimeout(function(){ document.getElementById("expenseAmount").focus(); }, 0);
  }
  function closeExpenseModal(){
    document.getElementById("expenseModalOverlay").hidden = true;
    editingExpenseId = null;
    receiptPhotoDraft = null;
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
    var payload = { clients:clients, tasks:tasks, meetings:meetings, projects:projects, expenses:expenses, monthlyValues:monthlyValues, exportedAt:new Date().toISOString() };
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
      var nMonthlyValues = Object.keys(parsed.monthlyValues||{}).length;
      var ok = window.confirm(
        "Importar " + nClients + " cliente(s), " + nTasks + " tarefa(s), " + nMeetings + " reunião(ões), " +
        nProjects + " projeto(s), " + nExpenses + " despesa(s) e " + nMonthlyValues + " valor(es) mensais?\n" +
        "Isto substitui todos os dados atuais deste navegador.\n" +
        "Nota: as fotos e documentos guardados neste dispositivo não fazem parte desta cópia e têm de ser voltados a anexar."
      );
      if (!ok) return;
      clients = parsed.clients || {};
      tasks = parsed.tasks || {};
      meetings = parsed.meetings || {};
      projects = parsed.projects || {};
      expenses = parsed.expenses || {};
      monthlyValues = parsed.monthlyValues || {};
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
    var labels = { home:"+ Nova tarefa", tasks:"+ Nova tarefa", meetings:"+ Nova reunião", projects:"+ Novo projeto", expenses:"+ Nova despesa", clients:"+ Nova tarefa" };
    document.getElementById("newItemBtn").textContent = labels[view] || "+ Novo";
    render();
  }

  // ---------- wiring ----------
  function wireStaticUI(){
    Array.prototype.forEach.call(document.querySelectorAll("#viewTabs button"), function(btn){
      btn.addEventListener("click", function(){ switchView(btn.dataset.view); });
    });

    document.getElementById("newItemBtn").addEventListener("click", function(){
      if (currentView === "home" || currentView === "tasks" || currentView === "clients") openTaskModal(null);
      else if (currentView === "meetings") openMeetingModal(null);
      else if (currentView === "projects") openProjectModal(null);
      else openExpenseModal(null);
    });
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

    ["prioritySeg","statusSeg","projectStatusSeg","expenseCategorySeg"].forEach(function(segId){
      document.getElementById(segId).addEventListener("click", function(e){
        var btn = e.target.closest("button");
        if (btn) setSegValue(segId, btn.dataset.val);
      });
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

    try { wireStaticUI(); }
    catch(e){ console.error("wireStaticUI falhou", e); showFatalError("ligar botões", e); }

    try { render(); }
    catch(e){ console.error("render falhou", e); showFatalError("mostrar conteúdo", e); }

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
      return JSON.parse(JSON.stringify({ clients:clients, tasks:tasks, meetings:meetings, projects:projects, expenses:expenses, monthlyValues:monthlyValues }));
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
      monthlyValues = data.monthlyValues || {};
      if (baseline){
        var target = { clients:clients, tasks:tasks, meetings:meetings, projects:projects, expenses:expenses, monthlyValues:monthlyValues };
        ["clients","tasks","meetings","projects","expenses","monthlyValues"].forEach(function(collection){
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
