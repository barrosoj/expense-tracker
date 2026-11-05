// Browser-side receipt OCR. The selected image is processed locally by Tesseract.js.
(function(){
  "use strict";

  function normalizeText(value){
    return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  }

  function parseAmount(value){
    var cleaned = String(value || "").replace(/EUR|€/gi, "").replace(/\s/g, "");
    if (cleaned.indexOf(",") >= 0){
      cleaned = cleaned.replace(/\./g, "").replace(",", ".");
    } else if ((cleaned.match(/\./g) || []).length > 1 || /\.\d{3}$/.test(cleaned)){
      cleaned = cleaned.replace(/\./g, "");
    }
    var amount = Number(cleaned);
    return Number.isFinite(amount) && amount > 0 ? amount : 0;
  }

  function extractAmount(lines){
    var candidates = [];
    var amountPattern = /(?:€\s*)?\d+(?:[.\s]\d{3})*(?:[,.]\d{2})|\d+(?:[,.]\d{2})?\s*€/g;
    lines.forEach(function(line, index){
      var normalized = normalizeText(line);
      if (!/\b(TOTAL|A PAGAR|VALOR|MONTANTE)\b/.test(normalized)) return;
      if (/\b(SUBTOTAL|IVA|DESCONTO|TROCO|ARTIGO|UNITARIO)\b/.test(normalized)) return;
      var rank = /A PAGAR|TOTAL A LIQUIDAR/.test(normalized) ? 3 : /TOTAL/.test(normalized) ? 2 : 1;
      var matches = line.match(amountPattern) || [];
      matches.forEach(function(match){
        var amount = parseAmount(match);
        if (amount > 0) candidates.push({ amount:amount, rank:rank, index:index });
      });
    });
    candidates.sort(function(a,b){ return b.rank - a.rank || b.index - a.index; });
    return candidates.length ? candidates[0].amount : null;
  }

  function extractMerchant(lines){
    var excluded = /\b(FATURA|RECIBO|TALAO|CONTRIBUINTE|NIF|TELEFONE|WWW|HTTP|DATA|HORA|DOCUMENTO|CAIXA|TERMINAL|OPERADOR|VENDA|CLIENTE|ARTIGO|QUANTIDADE|QTD|TOTAL|PAGAR|IVA|EUR)\b/;
    for (var i = 0; i < Math.min(lines.length, 8); i += 1){
      var line = String(lines[i] || "").replace(/\s+/g, " ").trim();
      var normalized = normalizeText(line);
      if (line.length >= 3 && /[A-Za-zÀ-ÿ]/.test(line) && !excluded.test(normalized)) return line.slice(0, 80);
    }
    return "";
  }

  function parseText(text){
    var lines = String(text || "").split(/\r?\n/).map(function(line){ return line.trim(); }).filter(Boolean);
    var supplier = extractMerchant(lines);
    return {
      amount: extractAmount(lines),
      supplier: supplier,
      description: supplier ? "Compra em " + supplier : ""
    };
  }

  function scan(file, onProgress){
    if (!file || !window.Tesseract || typeof window.Tesseract.createWorker !== "function"){
      return Promise.reject(new Error("Não foi possível carregar a leitura de recibos. Verifica a ligação à internet e tenta novamente."));
    }
    var worker;
    var timeout;
    var load = window.Tesseract.createWorker("por", 1, {
      cacheMethod: "write",
      langPath: "https://tessdata.projectnaptha.com/4.0.0_fast",
      logger: function(message){ if (typeof onProgress === "function") onProgress(message); }
    }).then(function(createdWorker){
      worker = createdWorker;
      return worker.recognize(file);
    }).then(function(result){
      var suggestion = parseText(result && result.data && result.data.text);
      suggestion.confidence = Number(result && result.data && result.data.confidence) || 0;
      if (!suggestion.amount && !suggestion.supplier){
        throw new Error("Não consegui identificar valor ou descrição. Podes preencher os campos manualmente.");
      }
      return suggestion;
    }).finally(function(){
      if (worker) worker.terminate();
    });
    var stalled = new Promise(function(resolve, reject){
      timeout = setTimeout(function(){
        reject(new Error("A leitura não avançou. Verifica a ligação à internet e tenta novamente."));
      }, 60000);
    });
    return Promise.race([load, stalled]).finally(function(){ clearTimeout(timeout); });
  }

  window.ReceiptScanner = { scan:scan, parseText:parseText };
})();
