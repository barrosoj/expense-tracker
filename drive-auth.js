// Return page of the Google Drive sign-in (see drive.js): keeps the access token on this device
// and goes back to the app. The token arrives in the URL fragment, which never reaches a server.
(function(){
  "use strict";
  var params = new URLSearchParams(location.hash.replace(/^#/, ""));
  var back = "./";
  try { back = localStorage.getItem("qt_drive_return") || "./"; localStorage.removeItem("qt_drive_return"); } catch(e){}
  // Only ever go back inside this app.
  try { if (new URL(back, location.href).origin !== location.origin) back = "./"; } catch(e){ back = "./"; }

  var expected = null;
  try { expected = localStorage.getItem("qt_drive_state"); localStorage.removeItem("qt_drive_state"); } catch(e){}
  var token = params.get("access_token");
  // Only accept a token from a sign-in this device started (blocks links that plant someone else's token).
  var stateOk = !!expected && params.get("state") === expected;
  history.replaceState(null, "", location.pathname);

  if (token && stateOk){
    var secs = parseInt(params.get("expires_in"), 10) || 3600;
    try { localStorage.setItem("qt_drive_token", JSON.stringify({ token: token, exp: Date.now() + secs * 1000 })); } catch(e){}
    location.replace(back);
    return;
  }
  var err = params.get("error");
  document.getElementById("msg").textContent = err === "access_denied"
    ? "Ligação ao Google Drive cancelada. A voltar à aplicação…"
    : "Não foi possível ligar ao Google Drive" + (err ? " (" + err + ")" : "") + ". A voltar à aplicação…";
  setTimeout(function(){ location.replace(back); }, 2000);
})();
