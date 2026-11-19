// Google Drive access: the Excel files (Compras, Vendas, Despesas) live in one Drive folder
// (APP_CONFIG.googleDriveFolderId) instead of on a PC. Sign-in uses Google's OAuth "token" flow
// with a full-page redirect to drive-auth.html (works in the iOS home-screen app, no library);
// the short-lived access token (about 1 hour) is kept on this device only.
(function(){
  "use strict";

  var cfg = window.APP_CONFIG || {};
  var TOKEN_KEY = "qt_drive_token";
  var RETURN_KEY = "qt_drive_return";
  var API = "https://www.googleapis.com/drive/v3/";
  var UPLOAD = "https://www.googleapis.com/upload/drive/v3/";
  var FOLDER_MIME = "application/vnd.google-apps.folder";
  var SHEET_MIME = "application/vnd.google-apps.spreadsheet";
  var XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  var SCOPE = "https://www.googleapis.com/auth/drive";

  function configured(){ return !!(cfg.googleClientId && cfg.googleDriveFolderId); }

  function readToken(){
    try {
      var t = JSON.parse(localStorage.getItem(TOKEN_KEY) || "null");
      // Treat it as expired a minute early so a request never starts with a dying token.
      if (t && t.token && t.exp > Date.now() + 60000) return t.token;
    } catch(e){}
    return null;
  }
  function connected(){ return configured() && !!readToken(); }
  function forget(){ try { localStorage.removeItem(TOKEN_KEY); } catch(e){} }

  function redirectUri(){ return new URL("drive-auth.html", location.href).href.split("#")[0].split("?")[0]; }

  // Leaves the app for Google's consent page; drive-auth.html stores the token and comes back here.
  function connect(){
    if (!configured()) return;
    var rnd = new Uint8Array(16);
    crypto.getRandomValues(rnd);
    var state = Array.prototype.map.call(rnd, function(b){ return ("0" + b.toString(16)).slice(-2); }).join("");
    try {
      localStorage.setItem("qt_drive_state", state);
      localStorage.setItem(RETURN_KEY, location.href.split("#")[0]);
    } catch(e){}
    var params = new URLSearchParams({
      client_id: cfg.googleClientId,
      redirect_uri: redirectUri(),
      response_type: "token",
      scope: SCOPE,
      include_granted_scopes: "true",
      state: state
    });
    location.href = "https://accounts.google.com/o/oauth2/v2/auth?" + params.toString();
  }

  function NotConnected(){ var e = new Error("sem ligação ao Google Drive"); e.notConnected = true; return e; }

  function api(url, options){
    var token = readToken();
    if (!token) return Promise.reject(NotConnected());
    options = options || {};
    options.headers = Object.assign({ Authorization: "Bearer " + token }, options.headers || {});
    return fetch(url, options).then(function(res){
      if (res.status === 401){ forget(); throw NotConnected(); }
      if (!res.ok){
        return res.json().catch(function(){ return {}; }).then(function(j){
          var msg = j && j.error && j.error.message;
          if (res.status === 404) msg = "pasta do Google Drive não encontrada ou sem acesso";
          throw new Error(msg || ("erro " + res.status + " no Google Drive"));
        });
      }
      return res;
    });
  }

  function q(s){ return "'" + String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'"; }

  // Children of a folder: [{ id, name, mimeType, modifiedTime }]
  function listChildren(folderId){
    var params = new URLSearchParams({
      q: q(folderId) + " in parents and trashed = false",
      fields: "files(id,name,mimeType,modifiedTime)",
      pageSize: "1000",
      orderBy: "name",
      supportsAllDrives: "true",
      includeItemsFromAllDrives: "true"
    });
    return api(API + "files?" + params.toString()).then(function(r){ return r.json(); }).then(function(j){ return j.files || []; });
  }

  // Sub-folder of the app's Drive folder by name (first match among `names`, case-insensitive).
  var folderCache = {};
  function findFolder(names, create){
    var key = names.join("|") + (create ? "+" : "");
    if (folderCache[key]) return folderCache[key];
    var p = listChildren(cfg.googleDriveFolderId).then(function(files){
      var folders = files.filter(function(f){ return f.mimeType === FOLDER_MIME; });
      for (var i = 0; i < names.length; i++){
        for (var k = 0; k < folders.length; k++){
          if (folders[k].name.toLowerCase() === names[i].toLowerCase()) return folders[k].id;
        }
      }
      if (!create) return null;
      return api(API + "files?supportsAllDrives=true&fields=id", {
        method:"POST",
        headers:{ "Content-Type":"application/json" },
        body: JSON.stringify({ name:names[0], mimeType:FOLDER_MIME, parents:[cfg.googleDriveFolderId] })
      }).then(function(r){ return r.json(); }).then(function(j){ return j.id; });
    });
    folderCache[key] = p;
    p.catch(function(){ delete folderCache[key]; });
    return p;
  }

  // File contents as an ArrayBuffer; native Google Sheets are exported as .xlsx.
  function download(file){
    var url = file.mimeType === SHEET_MIME
      ? API + "files/" + encodeURIComponent(file.id) + "/export?mimeType=" + encodeURIComponent(XLSX_MIME)
      : API + "files/" + encodeURIComponent(file.id) + "?alt=media&supportsAllDrives=true";
    return api(url, { cache:"no-store" }).then(function(r){ return r.arrayBuffer(); });
  }

  // Creates or replaces `name` in folder `folderId` (Drive keeps the old version in its history).
  function upload(folderId, name, bytes){
    return listChildren(folderId).then(function(files){
      var existing = files.filter(function(f){ return f.name === name && f.mimeType !== FOLDER_MIME; })[0];
      if (existing){
        return api(UPLOAD + "files/" + encodeURIComponent(existing.id) + "?uploadType=media&supportsAllDrives=true", {
          method:"PATCH", headers:{ "Content-Type":XLSX_MIME }, body:bytes
        });
      }
      var boundary = "qt" + Date.now().toString(36);
      var meta = JSON.stringify({ name:name, parents:[folderId], mimeType:XLSX_MIME });
      var body = new Blob([
        "--" + boundary + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + meta + "\r\n",
        "--" + boundary + "\r\nContent-Type: " + XLSX_MIME + "\r\n\r\n",
        bytes,
        "\r\n--" + boundary + "--"
      ]);
      return api(UPLOAD + "files?uploadType=multipart&supportsAllDrives=true", {
        method:"POST", headers:{ "Content-Type":"multipart/related; boundary=" + boundary }, body:body
      });
    });
  }

  window.QuadroDrive = {
    configured: configured,
    connected: connected,
    connect: connect,
    forget: forget,
    listChildren: listChildren,
    findFolder: findFolder,
    download: download,
    upload: upload,
    isSpreadsheet: function(f){ return f.mimeType === SHEET_MIME || /\.(xlsx|csv)$/i.test(f.name); }
  };
})();
