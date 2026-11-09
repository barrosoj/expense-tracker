// Offline-first sync to public.records. Each account only sees its own records (RLS).
// Pull/merge on sign-in and every 20s; local changes sync shortly after save.
(function(){
  "use strict";

  var COLLECTIONS = ["clients", "tasks", "meetings", "projects", "expenses", "monthlyValues"];
  var userId = null;
  var client = null;
  var running = false;
  var inFlight = false;
  var rerunRequested = false;
  var timer = null;
  var debounce = null;
  var statusEl = null;
  var lastSyncAt = 0;
  var RETRY_MS = 15000;
  var POLL_MS = 20000;

  function $(id){ return document.getElementById(id); }
  function localKey(){ return "qt_sync_shadow:" + userId; }
  function lastSyncKey(){ return "qt_last_sync_at:" + userId; }
  function showLastSync(timestamp){
    if (!statusEl) statusEl = $("syncStatus");
    if (!statusEl) return;
    statusEl.textContent = timestamp
      ? "Última sincronização: " + new Date(timestamp).toLocaleTimeString("pt-PT", { hour:"2-digit", minute:"2-digit" })
      : "Ainda não sincronizado";
  }
  function clone(value){ return JSON.parse(JSON.stringify(value)); }
  function blankData(){ return { clients:{}, tasks:{}, meetings:{}, projects:{}, expenses:{}, monthlyValues:{} }; }
  function getLocal(){
    var data = window.QuadroApp.getSyncData();
    COLLECTIONS.forEach(function(collection){
      Object.keys(data[collection] || {}).forEach(function(id){
        if (!itemTime(data[collection][id])) data[collection][id].updatedAt = Date.now();
      });
    });
    return data;
  }
  function itemTime(item){ return Number(item && item.updatedAt) || 0; }
  function key(collection, id){ return collection + "\u0000" + id; }
  function normalizeRemote(row){
    return {
      collection: row.collection,
      id: row.id,
      data: row.data || {},
      updatedAt: Number(row.updated_at) || itemTime(row.data),
      deletedAt: row.deleted_at == null ? null : Number(row.deleted_at)
    };
  }
  function readShadow(){
    try {
      var s = localStorage.getItem(localKey());
      return s ? JSON.parse(s) : blankData();
    } catch(e){ return blankData(); }
  }
  function writeShadow(data){
    try { localStorage.setItem(localKey(), JSON.stringify(data)); }
    catch(e){ console.warn("Could not save sync checkpoint", e); }
  }
  function readRemote(){
    var rows = [];
    function page(from){
      return client.from("records")
        .select("collection,id,data,updated_at,deleted_at")
        .eq("user_id", userId)
        .order("server_updated_at", { ascending:true })
        .order("collection", { ascending:true })
        .order("id", { ascending:true })
        .range(from, from + 999)
        .then(function(res){
          if (res.error) throw res.error;
          var batch = res.data || [];
          rows = rows.concat(batch);
          return batch.length === 1000 ? page(from + 1000) : rows;
        });
    }
    return page(0);
  }
  function makeRow(collection, id, item, deletedAt){
    var data = item ? clone(item) : {};
    var updatedAt = item ? (itemTime(item) || Date.now()) : deletedAt;
    if (item && !data.updatedAt) data.updatedAt = updatedAt;
    return {
      user_id: userId,
      collection: collection,
      id: id,
      data: data,
      updated_at: updatedAt,
      deleted_at: deletedAt == null ? null : deletedAt
    };
  }
  function upload(rows){
    if (!rows.length) return Promise.resolve();
    return client.from("records").upsert(rows, { onConflict:"user_id,collection,id" }).then(function(res){
      if (res.error) throw res.error;
    });
  }
  function mergeAndSync(){
    var local = getLocal();
    var shadow = readShadow();
    return readRemote().then(function(rawRows){
      var remote = {};
      rawRows.forEach(function(r){
        var normalized = normalizeRemote(r);
        remote[key(normalized.collection, normalized.id)] = normalized;
      });

      var merged = blankData();
      var outgoing = {};
      var now = Date.now();
      var allKeys = {};
      COLLECTIONS.forEach(function(c){
        Object.keys(local[c] || {}).forEach(function(id){ allKeys[key(c,id)] = { c:c, id:id }; });
        Object.keys(shadow[c] || {}).forEach(function(id){ allKeys[key(c,id)] = { c:c, id:id }; });
      });
      Object.keys(remote).forEach(function(k){
        if (!allKeys[k]){
          var split = k.split("\u0000");
          allKeys[k] = { c:split[0], id:split.slice(1).join("\u0000") };
        }
      });

      Object.keys(allKeys).forEach(function(k){
        var pair = allKeys[k], c = pair.c, id = pair.id;
        var localItem = local[c] && local[c][id];
        var shadowItem = shadow[c] && shadow[c][id];
        var remoteRow = remote[k];
        if (!remoteRow && !localItem && shadowItem){
          // Locally deleted since the last successful sync: retain a tombstone for other devices.
          outgoing[k] = makeRow(c, id, null, now);
          return;
        }
        if (remoteRow && remoteRow.deletedAt != null){
          if (localItem && itemTime(localItem) > remoteRow.deletedAt){
            merged[c][id] = localItem;
            outgoing[k] = makeRow(c, id, localItem, null);
          }
          return;
        }
        if (localItem && remoteRow){
          if (itemTime(localItem) > remoteRow.updatedAt){
            merged[c][id] = localItem;
            outgoing[k] = makeRow(c, id, localItem, null);
          } else {
            var remoteItem = clone(remoteRow.data);
            if (!remoteItem.updatedAt) remoteItem.updatedAt = remoteRow.updatedAt || now;
            merged[c][id] = remoteItem;
          }
          return;
        }
        if (localItem){
          merged[c][id] = localItem;
          outgoing[k] = makeRow(c, id, localItem, null);
          return;
        }
        if (remoteRow){
          // An item existed in the last synced copy and is now absent locally: treat as deletion.
          // New remote items (not in shadow) are downloaded normally.
          if (shadowItem){
            outgoing[k] = makeRow(c, id, null, now);
          } else {
            var downloaded = clone(remoteRow.data);
            if (!downloaded.updatedAt) downloaded.updatedAt = remoteRow.updatedAt || now;
            merged[c][id] = downloaded;
          }
        }
      });

      return upload(Object.keys(outgoing).map(function(k){ return outgoing[k]; })).then(function(){
        writeShadow(merged);
        if (window.QuadroApp.applySyncData(merged, local)) localChanged();
      });
    });
  }
  function syncNow(){
    if (!running) return Promise.resolve();
    if (inFlight){ rerunRequested = true; return Promise.resolve(); }
    if (!navigator.onLine) return Promise.resolve();
    inFlight = true;
    return mergeAndSync().then(function(){
      lastSyncAt = Date.now();
      try { localStorage.setItem(lastSyncKey(), String(lastSyncAt)); } catch(e){}
      showLastSync(lastSyncAt);
    }).catch(function(err){
      console.error("sync failed", err);
    }).then(function(){
      inFlight = false;
      if (rerunRequested){ rerunRequested = false; syncNow(); }
    });
  }
  function localChanged(){
    if (!running) return;
    clearTimeout(debounce);
    debounce = setTimeout(syncNow, 700);
  }
  function start(nextUserId, supabaseClient){
    if (running && userId === nextUserId) return;
    clearInterval(timer);
    clearTimeout(debounce);
    userId = nextUserId;
    client = supabaseClient;
    running = !!(userId && client);
    if (!running) return;
    statusEl = $("syncStatus");
    try { lastSyncAt = Number(localStorage.getItem(lastSyncKey())) || 0; } catch(e){ lastSyncAt = 0; }
    showLastSync(lastSyncAt);
    syncNow();
    timer = setInterval(syncNow, POLL_MS);
    window.addEventListener("online", syncNow);
    window.addEventListener("focus", syncNow);
    document.addEventListener("visibilitychange", function(){ if (!document.hidden) syncNow(); });
  }

  window.QuadroSync = { start:start, syncNow:syncNow, localChanged:localChanged };
})();
