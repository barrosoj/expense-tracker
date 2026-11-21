// Portuguese invoice QR code (Portaria 195/2020): every certified invoice carries a QR code with
// the seller's NIF, date, document number, ATCUD, VAT and total. Reading it is exact and works
// offline, unlike OCR. The decoder (vendor/jsqr.js, Apache-2.0) is loaded only when needed and
// the photo never leaves the device.
(function(){
  "use strict";

  var DOC_TYPES = {
    FT:"Fatura", FS:"Fatura simplificada", FR:"Fatura-recibo", NC:"Nota de crédito", ND:"Nota de débito",
    RG:"Recibo", GT:"Guia de transporte", GR:"Guia de remessa", OR:"Orçamento", PF:"Fatura pró-forma"
  };

  function nifOk(v){
    if (!/^\d{9}$/.test(v)) return false;
    var sum = 0;
    for (var i = 0; i < 8; i++) sum += Number(v.charAt(i)) * (9 - i);
    var check = 11 - (sum % 11);
    if (check >= 10) check = 0;
    return check === Number(v.charAt(8));
  }

  function money(v){
    var n = Number(String(v || "").replace(",", "."));
    return isFinite(n) ? Math.round(n * 100) / 100 : null;
  }

  // "A:123456789*B:999999990*...*O:12.30*..." → invoice fields, or null if it isn't an AT invoice QR.
  function parse(text){
    var raw = String(text || "").trim();
    if (!/^A:\d{9}\*/.test(raw)) return null;
    var f = {};
    raw.split("*").forEach(function(part){
      var i = part.indexOf(":");
      if (i > 0) f[part.slice(0, i).trim().toUpperCase()] = part.slice(i + 1).trim();
    });
    var nif = f.A || "";
    var d = /^(\d{4})(\d{2})(\d{2})$/.exec(f.F || "");
    var total = money(f.O);
    if (!d || total == null || total < 0) return null;
    var date = d[1] + "-" + d[2] + "-" + d[3];
    if (isNaN(new Date(date + "T00:00:00").getTime())) return null;
    var type = String(f.D || "").toUpperCase().slice(0, 2);
    return {
      nif: nif,
      nifValid: nifOk(nif),
      buyerNif: /^\d{9}$/.test(f.B || "") && f.B !== "999999990" ? f.B : "",
      type: type,
      typeLabel: DOC_TYPES[type] || "Documento",
      date: date,
      number: String(f.G || "").slice(0, 60),
      atcud: String(f.H || "").slice(0, 70),
      vat: money(f.N) || 0,
      total: total
    };
  }

  // ---------- decoding the photo ----------
  var loading = null;
  function loadDecoder(){
    if (window.jsQR) return Promise.resolve(window.jsQR);
    if (loading) return loading;
    loading = new Promise(function(resolve, reject){
      var s = document.createElement("script");
      s.src = "vendor/jsqr.js";
      s.onload = function(){ window.jsQR ? resolve(window.jsQR) : reject(new Error("leitor de QR indisponível")); };
      s.onerror = function(){ loading = null; reject(new Error("leitor de QR indisponível")); };
      document.head.appendChild(s);
    });
    return loading;
  }

  function loadImage(file){
    if (window.createImageBitmap){
      return createImageBitmap(file).catch(function(){ return loadImgTag(file); });
    }
    return loadImgTag(file);
  }
  function loadImgTag(file){
    return new Promise(function(resolve, reject){
      var url = URL.createObjectURL(file);
      var img = new Image();
      img.onload = function(){ URL.revokeObjectURL(url); resolve(img); };
      img.onerror = function(){ URL.revokeObjectURL(url); reject(new Error("imagem inválida")); };
      img.src = url;
    });
  }

  // Tries the whole photo at two sizes, then each half/quarter at a higher zoom: the QR is often
  // small at the bottom of a long receipt.
  function decode(jsQR, img){
    var W = img.width, H = img.height;
    var regions = [
      [0, 0, 1, 1, 1400], [0, 0, 1, 1, 900],
      [0, 0.5, 1, 0.5, 1400], [0, 0.66, 1, 0.34, 1200], [0, 0, 1, 0.5, 1400],
      [0, 0.5, 0.5, 0.5, 1200], [0.5, 0.5, 0.5, 0.5, 1200], [0.25, 0.4, 0.5, 0.6, 1200]
    ];
    var canvas = document.createElement("canvas");
    var ctx = canvas.getContext("2d", { willReadFrequently:true });
    for (var i = 0; i < regions.length; i++){
      var r = regions[i];
      var sx = r[0] * W, sy = r[1] * H, sw = r[2] * W, sh = r[3] * H;
      var scale = Math.min(1, r[4] / Math.max(sw, sh));
      var cw = Math.max(1, Math.round(sw * scale)), ch = Math.max(1, Math.round(sh * scale));
      canvas.width = cw; canvas.height = ch;
      ctx.drawImage(img, sx, sy, sw, sh, 0, 0, cw, ch);
      var data = ctx.getImageData(0, 0, cw, ch);
      var code = jsQR(data.data, cw, ch, { inversionAttempts: i < 2 ? "attemptBoth" : "dontInvert" });
      if (code && code.data){
        var inv = parse(code.data);
        if (inv) return inv;
      }
    }
    return null;
  }

  // Resolves with the invoice fields, or null when the photo has no readable AT QR code.
  function scan(file){
    return Promise.all([loadDecoder(), loadImage(file)]).then(function(res){
      var img = res[1];
      try { return decode(res[0], img); }
      finally { if (img.close) img.close(); }
    });
  }

  window.QuadroInvoiceQR = { parse: parse, scan: scan, nifOk: nifOk };
})();
