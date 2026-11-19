// Minimal .xlsx reader: returns the cells of the first worksheet as rows of values.
// Runs entirely in the browser — the file is never uploaded anywhere.
// An .xlsx file is a zip of XML parts; entries are unzipped with the browser's
// built-in DecompressionStream("deflate-raw"), so no external library is needed.
(function(){
  "use strict";

  function u16(b, o){ return b[o] | (b[o+1] << 8); }
  function u32(b, o){ return (b[o] | (b[o+1] << 8) | (b[o+2] << 16) | (b[o+3] << 24)) >>> 0; }

  function readZipEntries(bytes){
    // End of central directory record: scan backwards for its signature.
    var eocd = -1;
    for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--){
      if (u32(bytes, i) === 0x06054b50){ eocd = i; break; }
    }
    if (eocd < 0) throw new Error("O ficheiro não é um Excel (.xlsx) válido.");
    var count = u16(bytes, eocd + 10);
    var p = u32(bytes, eocd + 16);
    var entries = {};
    var decoder = new TextDecoder("utf-8");
    for (var n = 0; n < count; n++){
      if (u32(bytes, p) !== 0x02014b50) throw new Error("Estrutura do ficheiro Excel inválida.");
      var method = u16(bytes, p + 10);
      var compSize = u32(bytes, p + 20);
      var nameLen = u16(bytes, p + 28);
      var extraLen = u16(bytes, p + 30);
      var commentLen = u16(bytes, p + 32);
      var localOffset = u32(bytes, p + 42);
      var name = decoder.decode(bytes.subarray(p + 46, p + 46 + nameLen));
      entries[name] = { method:method, compSize:compSize, localOffset:localOffset };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  function readEntryText(bytes, entry){
    var lo = entry.localOffset;
    var start = lo + 30 + u16(bytes, lo + 26) + u16(bytes, lo + 28);
    var data = bytes.subarray(start, start + entry.compSize);
    if (entry.method === 0) return Promise.resolve(new TextDecoder("utf-8").decode(data));
    if (entry.method !== 8) return Promise.reject(new Error("Compressão do Excel não suportada."));
    if (typeof DecompressionStream === "undefined"){
      return Promise.reject(new Error("Este navegador não consegue ler ficheiros Excel. Atualiza-o ou usa Chrome/Edge."));
    }
    var stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Response(stream).text();
  }

  function parseXml(text){
    var doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("XML inválido dentro do Excel.");
    return doc;
  }
  function byTag(node, tag){ return node.getElementsByTagNameNS("*", tag); }

  // "B12" → 1 (zero-based column index)
  function columnIndex(ref){
    var letters = String(ref || "").replace(/[^A-Z]/gi, "").toUpperCase();
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }

  function firstSheetPath(entries, workbookXml, relsXml){
    var fallback = "xl/worksheets/sheet1.xml";
    try {
      var sheet = byTag(parseXml(workbookXml), "sheet")[0];
      if (!sheet) return fallback;
      var rid = sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || sheet.getAttribute("r:id");
      var rels = byTag(parseXml(relsXml), "Relationship");
      for (var i = 0; i < rels.length; i++){
        if (rels[i].getAttribute("Id") === rid){
          var target = rels[i].getAttribute("Target").replace(/^\//, "");
          var path = target.indexOf("xl/") === 0 ? target : "xl/" + target;
          if (entries[path]) return path;
        }
      }
    } catch(e){ /* use fallback */ }
    return fallback;
  }

  // Resolves to { sheetName, rows } where rows[r][c] is a string, a number or null.
  function readXlsx(arrayBuffer){
    var bytes = new Uint8Array(arrayBuffer);
    var entries;
    try { entries = readZipEntries(bytes); }
    catch(e){ return Promise.reject(e); }
    function text(name){ return entries[name] ? readEntryText(bytes, entries[name]) : Promise.resolve(null); }

    return Promise.all([text("xl/workbook.xml"), text("xl/_rels/workbook.xml.rels"), text("xl/sharedStrings.xml")])
      .then(function(parts){
        var workbookXml = parts[0], relsXml = parts[1], sharedXml = parts[2];
        if (!workbookXml) throw new Error("O ficheiro não é um Excel (.xlsx) válido.");
        var sheetName = "";
        try { sheetName = byTag(parseXml(workbookXml), "sheet")[0].getAttribute("name") || ""; } catch(e){}
        var shared = [];
        if (sharedXml){
          var si = byTag(parseXml(sharedXml), "si");
          for (var i = 0; i < si.length; i++){
            var ts = byTag(si[i], "t"), s = "";
            for (var j = 0; j < ts.length; j++) s += ts[j].textContent;
            shared.push(s);
          }
        }
        var path = relsXml ? firstSheetPath(entries, workbookXml, relsXml) : "xl/worksheets/sheet1.xml";
        return text(path).then(function(sheetXml){
          if (!sheetXml) throw new Error("O Excel não tem nenhuma folha.");
          var rows = [];
          var rowEls = byTag(parseXml(sheetXml), "row");
          for (var r = 0; r < rowEls.length; r++){
            var rowIdx = parseInt(rowEls[r].getAttribute("r"), 10);
            rowIdx = isNaN(rowIdx) ? rows.length : rowIdx - 1;
            var row = [];
            var cells = byTag(rowEls[r], "c");
            for (var c = 0; c < cells.length; c++){
              var cell = cells[c];
              var col = cell.getAttribute("r") ? columnIndex(cell.getAttribute("r")) : row.length;
              var type = cell.getAttribute("t");
              var vEl = byTag(cell, "v")[0];
              var raw = vEl ? vEl.textContent : null;
              var value = null;
              if (type === "s") value = raw !== null ? (shared[parseInt(raw, 10)] || "") : null;
              else if (type === "inlineStr"){
                var t = byTag(cell, "t"), s2 = "";
                for (var k = 0; k < t.length; k++) s2 += t[k].textContent;
                value = s2;
              }
              else if (type === "str" || type === "e") value = raw;
              else if (type === "b") value = raw === "1";
              else if (raw !== null && raw !== ""){
                var num = Number(raw);
                value = isNaN(num) ? raw : num;
              }
              if (typeof value === "string") value = value.trim();
              row[col] = value;
            }
            rows[rowIdx] = row;
          }
          for (var x = 0; x < rows.length; x++) if (!rows[x]) rows[x] = [];
          return { sheetName:sheetName, rows:rows };
        });
      });
  }

  window.QuadroXlsx = { read: readXlsx };
})();
