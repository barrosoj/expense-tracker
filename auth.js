// Authentication gate (Supabase Auth): email one-time code + OAuth providers.
// Runs before the app. If Supabase is not configured, starts the app in local-only mode.
(function(){
  "use strict";

  var cfg = window.APP_CONFIG || {};
  var PROVIDER_LABEL = { google:"Google", apple:"Apple", github:"GitHub" };

  var appStarted = false;
  var pendingEmail = "";
  var resendTimer = null;
  var activeUserId = null;
  var sessionTimer = null;
  var MAX_SESSION_AGE_MS = 2 * 60 * 60 * 1000;
  var EXPIRED_NOTICE_KEY = "qt_session_expired_notice";

  function $(id){ return document.getElementById(id); }

  // userId: signed-in user (each user gets separate local data); omitted in local-only mode.
  function startApp(userId){
    document.querySelector(".app").hidden = false;
    if (appStarted) return;
    appStarted = true;
    if (window.QuadroApp && typeof window.QuadroApp.start === "function") window.QuadroApp.start({ userId: userId || null });
  }

  function isConfigured(){
    return !!(cfg.supabaseUrl && cfg.supabaseAnonKey);
  }

  // Local-only mode: no backend configured.
  if (!isConfigured()){
    startApp();
    return;
  }

  if (!window.supabase || typeof window.supabase.createClient !== "function"){
    console.error("supabase-js não carregou");
    startApp(); // degrade gracefully to local mode rather than locking the user out
    return;
  }

  var client = window.supabase.createClient(cfg.supabaseUrl, cfg.supabaseAnonKey, {
    auth: {
      flowType: "pkce",
      persistSession: true,
      autoRefreshToken: true,
      detectSessionInUrl: true
    }
  });
  window.QuadroAuth = { client: client, signOut: signOut };

  // ---------- UI helpers ----------
  function showScreen(){ $("authScreen").hidden = false; document.querySelector(".app").hidden = true; }
  function hideScreen(){ $("authScreen").hidden = true; document.querySelector(".app").hidden = false; }

  function showStep(step){
    $("authStepEmail").hidden = step !== "email";
    $("authStepCode").hidden = step !== "code";
    setMessage("");
    setTimeout(function(){
      var focusEl = step === "email" ? $("authEmail") : $("authCode");
      if (focusEl) focusEl.focus();
    }, 0);
  }

  function setMessage(text, isError){
    var m = $("authMessage");
    m.textContent = text || "";
    m.classList.toggle("error", !!isError);
  }

  function setBusy(busy){
    Array.prototype.forEach.call(document.querySelectorAll("#authScreen button, #authScreen input"), function(el){
      el.disabled = busy;
    });
  }

  function friendlyError(err){
    var msg = (err && (err.message || err.error_description)) || String(err || "");
    var code = err && err.code;
    if (code === "otp_expired" || /expired|invalid/i.test(msg) && /otp|token/i.test(msg)) return "Código inválido ou expirado. Pede um novo código.";
    if (code === "over_email_send_rate_limit" || /rate limit|too many/i.test(msg)) return "Demasiados pedidos. Espera um minuto e tenta novamente.";
    if (code === "otp_disabled" || code === "signup_disabled" || /signups not allowed|not allowed for otp/i.test(msg)) return "Este email não tem acesso a esta aplicação.";
    if (/provider is not enabled|unsupported provider/i.test(msg)) return "Este método de entrada não está ativo.";
    if (/failed to fetch|network/i.test(msg)) return "Sem ligação à internet. Tenta novamente quando estiveres online.";
    return msg || "Ocorreu um erro. Tenta novamente.";
  }

  function isIOSStandalone(){
    return window.navigator.standalone === true ||
      (window.matchMedia && window.matchMedia("(display-mode: standalone)").matches && /iPad|iPhone|iPod/.test(navigator.userAgent));
  }

  function startResendCooldown(seconds){
    var btn = $("authResendBtn");
    var left = seconds;
    clearInterval(resendTimer);
    btn.disabled = true;
    btn.textContent = "Reenviar código (" + left + "s)";
    resendTimer = setInterval(function(){
      left -= 1;
      if (left <= 0){
        clearInterval(resendTimer);
        btn.disabled = false;
        btn.textContent = "Reenviar código";
      } else {
        btn.textContent = "Reenviar código (" + left + "s)";
      }
    }, 1000);
  }

  // ---------- email code ----------
  function sendCode(email){
    setBusy(true);
    setMessage("A enviar código...");
    return client.auth.signInWithOtp({
      email: email,
      options: {
        shouldCreateUser: !!cfg.allowEmailSignUp,
        // Used only if the email contains a link instead of a code (default Supabase template):
        // opening it in this same browser signs in here.
        emailRedirectTo: location.origin + location.pathname
      }
    }).then(function(res){
      setBusy(false);
      if (res.error) throw res.error;
      pendingEmail = email;
      $("authCodeEmail").textContent = email;
      $("authCode").value = "";
      showStep("code");
      setMessage("Enviámos um código para o teu email (pode demorar um minuto). Se receberes um link em vez de código, abre-o neste mesmo navegador.");
      startResendCooldown(60);
    }).catch(function(err){
      setBusy(false);
      setMessage(friendlyError(err), true);
    });
  }

  function verifyCode(code){
    setBusy(true);
    setMessage("A verificar...");
    return client.auth.verifyOtp({ email: pendingEmail, token: code, type: "email" }).then(function(res){
      setBusy(false);
      if (res.error) throw res.error;
      // onAuthStateChange handles the transition into the app
    }).catch(function(err){
      setBusy(false);
      setMessage(friendlyError(err), true);
      $("authCode").select();
    });
  }

  // ---------- OAuth providers ----------
  function signInWithProvider(provider){
    setBusy(true);
    setMessage("A redirecionar para " + (PROVIDER_LABEL[provider] || provider) + "...");
    client.auth.signInWithOAuth({
      provider: provider,
      options: { redirectTo: location.origin + location.pathname }
    }).then(function(res){
      if (res.error) throw res.error;
      // browser navigates away to the provider
    }).catch(function(err){
      setBusy(false);
      setMessage(friendlyError(err), true);
    });
  }

  function renderProviders(){
    var wrap = $("authProviders");
    wrap.innerHTML = "";
    var providers = (cfg.authProviders || []).filter(function(p){ return PROVIDER_LABEL[p]; });
    // emailLogin:false → providers only (no email code form). Never hide everything.
    var emailLogin = cfg.emailLogin !== false || providers.length === 0;
    $("authDivider").hidden = providers.length === 0 || !emailLogin;
    $("authEmailForm").hidden = !emailLogin;
    providers.forEach(function(p){
      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn " + (emailLogin ? "btn-ghost" : "btn-primary") + " auth-provider auth-provider-" + p;
      btn.textContent = "Continuar com " + PROVIDER_LABEL[p];
      btn.addEventListener("click", function(){ signInWithProvider(p); });
      wrap.appendChild(btn);
    });
    // On iOS home-screen apps, OAuth may finish in Safari (separate storage), leaving
    // the session there instead of in the app. Show a hint with the fallback.
    var hint = $("authIOSHint");
    hint.textContent = emailLogin
      ? "No ecrã principal do iPhone/iPad, usa o código por email — a entrada com Google pode terminar no Safari e a sessão ficar lá."
      : "No ecrã principal do iPhone/iPad: se depois de entrar com Google voltares ao Safari em vez da app, usa a app diretamente no Safari.";
    hint.hidden = !(providers.length && isIOSStandalone());
  }

  // ---------- session ----------
  // scope "local"  = this device only.
  // scope "global" = revoke every session of this account (other devices are signed out
  //                  when their current access token expires, at most ~1 hour).
  function signOut(scope){
    scope = scope === "global" ? "global" : "local";
    return client.auth.signOut({ scope: scope }).then(function(res){
      if (res && res.error) throw res.error;
    }).catch(function(err){
      console.error("signOut failed", err);
      if (scope === "global") alert("Não foi possível terminar as sessões nos outros dispositivos. Verifica a ligação à internet e tenta novamente.");
    });
  }

  // Last signed-in user on this device: { id, email }. Used to open the right local data offline.
  var LAST_USER_KEY = "qt_last_signed_in_user";

  function sessionStartKey(userId){ return "qt_session_started_at:" + userId; }
  function sessionStartedAt(userId){
    var startedAt = 0;
    try { startedAt = Number(localStorage.getItem(sessionStartKey(userId))) || 0; } catch(e){}
    if (!startedAt){
      startedAt = Date.now();
      try { localStorage.setItem(sessionStartKey(userId), String(startedAt)); } catch(e){}
    }
    return startedAt;
  }
  function stopSessionTimer(){ clearTimeout(sessionTimer); sessionTimer = null; }
  function expireSession(){
    if (!activeUserId) return;
    stopSessionTimer();
    try { localStorage.setItem(EXPIRED_NOTICE_KEY, "1"); } catch(e){}
    signOut("local");
  }
  function checkSessionAge(){
    if (!activeUserId) return false;
    var remaining = MAX_SESSION_AGE_MS - (Date.now() - sessionStartedAt(activeUserId));
    if (remaining <= 0){ expireSession(); return true; }
    stopSessionTimer();
    sessionTimer = setTimeout(expireSession, remaining);
    return false;
  }

  function showSignedIn(session){
    var user = (session && session.user) || {};
    if (!user.id) return;
    activeUserId = user.id;
    if (checkSessionAge()) return;
    try { localStorage.setItem(LAST_USER_KEY, JSON.stringify({ id: user.id, email: user.email || "" })); } catch(e){}
    $("accountEmail").textContent = user.email || "";
    $("accountBox").hidden = false;
    hideScreen();
    startApp(user.id);
    if (window.QuadroSync) window.QuadroSync.start(user.id, client);
  }

  // Offline-first: if this device was signed in before and we're offline (token refresh
  // impossible), let that same user keep working on their local copy instead of locking them out.
  function tryOfflineEntry(){
    var last = null;
    try { last = JSON.parse(localStorage.getItem(LAST_USER_KEY) || "null"); } catch(e){}
    if (!last || !last.id || navigator.onLine !== false) return false;
    activeUserId = last.id;
    if (checkSessionAge()) return false;
    $("accountEmail").textContent = (last.email ? last.email + " · " : "") + "offline";
    $("accountBox").hidden = false;
    hideScreen();
    startApp(last.id);
    if (window.QuadroSync) window.QuadroSync.start(last.id, client);
    return true;
  }

  function showSignedOut(){
    stopSessionTimer();
    if (activeUserId){
      try { localStorage.removeItem(sessionStartKey(activeUserId)); } catch(e){}
    }
    activeUserId = null;
    try { localStorage.removeItem(LAST_USER_KEY); } catch(e){}
    $("accountBox").hidden = true;
    pendingEmail = "";
    showScreen();
    showStep("email");
    try {
      if (localStorage.getItem(EXPIRED_NOTICE_KEY)){
        localStorage.removeItem(EXPIRED_NOTICE_KEY);
        setMessage("Por segurança, a sessão terminou após 2 horas. Volta a entrar.");
      }
    } catch(e){}
  }

  function cleanAuthParamsFromUrl(){
    if (/[?&#](code|error|error_description|access_token)=/.test(location.href)){
      history.replaceState(null, "", location.pathname);
    }
  }

  function readOAuthErrorFromUrl(){
    var params = new URLSearchParams(location.search + "&" + location.hash.replace(/^#/, ""));
    var desc = params.get("error_description") || params.get("error");
    return desc ? desc.replace(/\+/g, " ") : null;
  }

  // ---------- wiring ----------
  function wire(){
    $("authEmailForm").addEventListener("submit", function(e){
      e.preventDefault();
      var email = $("authEmail").value.trim().toLowerCase();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){ setMessage("Indica um email válido.", true); return; }
      sendCode(email);
    });
    $("authCodeForm").addEventListener("submit", function(e){
      e.preventDefault();
      var code = $("authCode").value.replace(/\D/g, "");
      if (code.length < 6){ setMessage("O código tem pelo menos 6 dígitos.", true); return; }
      verifyCode(code);
    });
    $("authCode").addEventListener("input", function(){
      this.value = this.value.replace(/\D/g, "").slice(0, 10);
    });
    $("authResendBtn").addEventListener("click", function(){ if (pendingEmail) sendCode(pendingEmail); });
    $("authBackBtn").addEventListener("click", function(){ clearInterval(resendTimer); showStep("email"); });
    $("signOutBtn").addEventListener("click", function(){
      if (window.confirm("Terminar sessão neste dispositivo?")) signOut("local");
    });
    $("syncNowBtn").addEventListener("click", function(){
      if (window.QuadroSync) window.QuadroSync.syncNow();
    });
    $("signOutAllBtn").addEventListener("click", function(){
      if (window.confirm(
        "Terminar a sessão desta conta em TODOS os dispositivos (PC, telemóvel, tablet)?\n\n" +
        "Os outros dispositivos saem em até 1 hora. Os dados guardados neles não são apagados."
      )) signOut("global");
    });
    window.addEventListener("focus", checkSessionAge);
    document.addEventListener("visibilitychange", function(){
      if (!document.hidden) checkSessionAge();
    });
    renderProviders();
  }

  function init(){
    wire();

    var oauthError = readOAuthErrorFromUrl();

    client.auth.onAuthStateChange(function(event, session){
      // Defer to avoid running app code inside the auth callback lock.
      setTimeout(function(){
        if (session) showSignedIn(session);
        else if (event === "SIGNED_OUT"){
          // Reload to drop in-memory state from the previous session.
          location.reload();
        }
      }, 0);
    });

    client.auth.getSession().then(function(res){
      cleanAuthParamsFromUrl();
      var session = res && res.data && res.data.session;
      if (session){ showSignedIn(session); return; }
      if (tryOfflineEntry()) return;
      showSignedOut();
      if (oauthError) setMessage(friendlyError({ message: oauthError }), true);
      else if (res && res.error) setMessage(friendlyError(res.error), true);
    }).catch(function(err){
      cleanAuthParamsFromUrl();
      if (tryOfflineEntry()) return;
      showSignedOut();
      setMessage(friendlyError(err), true);
    });
  }

  init();
})();
