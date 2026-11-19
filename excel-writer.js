// Fills copies of the user's Excel templates (Despesas.xlsx, Combustiveis.xlsx) with expenses.
// Runs in the browser. Only the worksheet XML of the copy is changed: every other part of the
// file (styles, formulas, images, comments, print settings…) is copied byte for byte, so the
// result keeps the original format. Nothing here sends data anywhere.
(function(){
  "use strict";

  var XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
  var NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  var NS_XML = "http://www.w3.org/XML/1998/namespace";

  // ---------- zip (read + rewrite) ----------
  function u16(b, o){ return b[o] | (b[o+1] << 8); }
  function u32(b, o){ return (b[o] | (b[o+1] << 8) | (b[o+2] << 16) | (b[o+3] << 24)) >>> 0; }

  var CRC_TABLE = (function(){
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++){
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes){
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function readZip(bytes){
    var eocd = -1;
    for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--){
      if (u32(bytes, i) === 0x06054b50){ eocd = i; break; }
    }
    if (eocd < 0) throw new Error("o modelo não é um ficheiro Excel (.xlsx) válido");
    var count = u16(bytes, eocd + 10);
    var p = u32(bytes, eocd + 16);
    var entries = [];
    for (var n = 0; n < count; n++){
      if (u32(bytes, p) !== 0x02014b50) throw new Error("estrutura do Excel modelo inválida");
      var nameLen = u16(bytes, p + 28), extraLen = u16(bytes, p + 30), commentLen = u16(bytes, p + 32);
      var e = {
        flags: u16(bytes, p + 8), method: u16(bytes, p + 10),
        time: u16(bytes, p + 12), date: u16(bytes, p + 14),
        crc: u32(bytes, p + 16), compSize: u32(bytes, p + 20), size: u32(bytes, p + 24),
        nameBytes: bytes.slice(p + 46, p + 46 + nameLen)
      };
      e.name = new TextDecoder("utf-8").decode(e.nameBytes);
      var lo = u32(bytes, p + 42);
      var start = lo + 30 + u16(bytes, lo + 26) + u16(bytes, lo + 28);
      e.data = bytes.subarray(start, start + e.compSize);
      entries.push(e);
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  function streamBytes(data, stream){
    return new Response(new Blob([data]).stream().pipeThrough(stream)).arrayBuffer().then(function(b){ return new Uint8Array(b); });
  }
  function inflate(entry){
    if (entry.method === 0) return Promise.resolve(entry.data);
    if (entry.method !== 8) return Promise.reject(new Error("compressão do Excel não suportada"));
    if (typeof DecompressionStream === "undefined") return Promise.reject(new Error("este navegador não consegue ler Excel"));
    return streamBytes(entry.data, new DecompressionStream("deflate-raw"));
  }
  function deflate(bytes){
    if (typeof CompressionStream === "undefined") return Promise.resolve(null);
    return streamBytes(bytes, new CompressionStream("deflate-raw")).catch(function(){ return null; });
  }

  function writeZip(entries){
    var parts = [], central = [], offset = 0;
    function header(size){ var b = new Uint8Array(size); return { b:b, v:new DataView(b.buffer) }; }
    entries.forEach(function(e){
      var flags = e.flags & 0x0800; // keep "UTF-8 name"; sizes are written up front (no data descriptor)
      var h = header(30);
      h.v.setUint32(0, 0x04034b50, true); h.v.setUint16(4, 20, true); h.v.setUint16(6, flags, true);
      h.v.setUint16(8, e.method, true); h.v.setUint16(10, e.time, true); h.v.setUint16(12, e.date, true);
      h.v.setUint32(14, e.crc, true); h.v.setUint32(18, e.data.length, true); h.v.setUint32(22, e.size, true);
      h.v.setUint16(26, e.nameBytes.length, true); h.v.setUint16(28, 0, true);
      parts.push(h.b, e.nameBytes, e.data);
      var c = header(46);
      c.v.setUint32(0, 0x02014b50, true); c.v.setUint16(4, 20, true); c.v.setUint16(6, 20, true);
      c.v.setUint16(8, flags, true); c.v.setUint16(10, e.method, true); c.v.setUint16(12, e.time, true);
      c.v.setUint16(14, e.date, true); c.v.setUint32(16, e.crc, true); c.v.setUint32(20, e.data.length, true);
      c.v.setUint32(24, e.size, true); c.v.setUint16(28, e.nameBytes.length, true);
      c.v.setUint32(42, offset, true);
      central.push(c.b, e.nameBytes);
      offset += 30 + e.nameBytes.length + e.data.length;
    });
    var centralSize = central.reduce(function(s, b){ return s + b.length; }, 0);
    var end = header(22);
    end.v.setUint32(0, 0x06054b50, true);
    end.v.setUint16(8, entries.length, true); end.v.setUint16(10, entries.length, true);
    end.v.setUint32(12, centralSize, true); end.v.setUint32(16, offset, true);
    var all = parts.concat(central, [end.b]);
    var out = new Uint8Array(all.reduce(function(s, b){ return s + b.length; }, 0));
    var pos = 0;
    all.forEach(function(b){ out.set(b, pos); pos += b.length; });
    return out;
  }

  function Package(buffer){
    this.entries = readZip(new Uint8Array(buffer));
    this.replaced = {};
    this.added = [];   // names of new parts
    this.removed = {}; // names of parts to leave out
  }
  Package.prototype.find = function(name){
    for (var i = 0; i < this.entries.length; i++) if (this.entries[i].name === name) return this.entries[i];
    return null;
  };
  Package.prototype.text = function(name){
    var e = this.find(name);
    if (!e) return Promise.resolve(null);
    return inflate(e).then(function(b){ return new TextDecoder("utf-8").decode(b); });
  };
  Package.prototype.has = function(name){
    return !this.removed[name] && (!!this.find(name) || this.added.indexOf(name) >= 0);
  };
  Package.prototype.setText = function(name, text){
    delete this.removed[name];
    if (!this.find(name) && this.added.indexOf(name) < 0) this.added.push(name);
    this.replaced[name] = new TextEncoder().encode(text);
  };
  Package.prototype.remove = function(name){
    this.removed[name] = true;
    delete this.replaced[name];
    this.added = this.added.filter(function(n){ return n !== name; });
  };
  Package.prototype.build = function(){
    var self = this;
    var now = new Date();
    var dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    var dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    var all = this.entries.filter(function(e){ return !self.removed[e.name]; }).concat(this.added.map(function(name){
      return { name:name, nameBytes:new TextEncoder().encode(name), flags:0, time:dosTime, date:dosDate };
    }));
    return Promise.all(all.map(function(e){
      var raw = self.replaced[e.name];
      if (!raw) return Promise.resolve(e);
      return deflate(raw).then(function(comp){
        return {
          name:e.name, nameBytes:e.nameBytes, flags:e.flags, time:e.time, date:e.date,
          method: comp ? 8 : 0, data: comp || raw, crc: crc32(raw), size: raw.length
        };
      });
    })).then(writeZip);
  };

  // ---------- workbook / worksheet helpers ----------
  function parseXml(text){
    var doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("XML inválido no modelo");
    return doc;
  }
  function serializeXml(doc){
    return XML_DECL + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, "");
  }
  function byTag(node, tag){ return Array.prototype.slice.call(node.getElementsByTagNameNS("*", tag)); }
  function plain(s){
    return String(s == null ? "" : s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z0-9]/g, "");
  }

  // Sheets of the workbook in order: [{ name, path }]
  function listSheets(pkg){
    return Promise.all([pkg.text("xl/workbook.xml"), pkg.text("xl/_rels/workbook.xml.rels"), pkg.text("xl/sharedStrings.xml")])
      .then(function(parts){
        var wb = parseXml(parts[0]);
        var targets = {};
        byTag(parseXml(parts[1]), "Relationship").forEach(function(r){ targets[r.getAttribute("Id")] = r.getAttribute("Target"); });
        var shared = [];
        if (parts[2]){
          byTag(parseXml(parts[2]), "si").forEach(function(si){
            shared.push(byTag(si, "t").map(function(t){ return t.textContent; }).join(""));
          });
        }
        var sheets = byTag(wb, "sheet").map(function(s){
          var t = (targets[s.getAttributeNS(NS_REL, "id") || s.getAttribute("r:id")] || "").replace(/^\//, "");
          return { name:s.getAttribute("name"), path: t.indexOf("xl/") === 0 ? t : "xl/" + t };
        });
        return { workbookDoc:wb, sheets:sheets, shared:shared };
      });
  }

  function colIndex(letters){
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }
  function splitRef(ref){
    var m = /^([A-Z]+)(\d+)$/.exec(ref);
    return { col:m[1], row:parseInt(m[2], 10) };
  }

  function Sheet(doc, shared){
    this.doc = doc;
    this.shared = shared;
    this.ns = doc.documentElement.namespaceURI;
    this.data = byTag(doc, "sheetData")[0];
    this.rows = {};
    var self = this;
    byTag(this.data, "row").forEach(function(r){ self.rows[parseInt(r.getAttribute("r"), 10)] = r; });
  }
  Sheet.prototype.row = function(n, create){
    if (this.rows[n] || !create) return this.rows[n] || null;
    var r = this.doc.createElementNS(this.ns, "row");
    r.setAttribute("r", String(n));
    var next = null;
    Object.keys(this.rows).map(Number).sort(function(a, b){ return a - b; }).some(function(k){
      if (k > n){ next = k; return true; }
      return false;
    });
    this.data.insertBefore(r, next ? this.rows[next] : null);
    this.rows[n] = r;
    return r;
  };
  Sheet.prototype.cell = function(ref, create){
    var p = splitRef(ref);
    var row = this.row(p.row, create);
    if (!row) return null;
    var cells = byTag(row, "c"), target = colIndex(p.col), before = null;
    for (var i = 0; i < cells.length; i++){
      var r = cells[i].getAttribute("r");
      if (r === ref) return cells[i];
      if (!before && r && colIndex(splitRef(r).col) > target) before = cells[i];
    }
    if (!create) return null;
    var c = this.doc.createElementNS(this.ns, "c");
    c.setAttribute("r", ref);
    row.insertBefore(c, before);
    return c;
  };
  Sheet.prototype.styleOf = function(ref){
    var c = this.cell(ref, false);
    return c && c.getAttribute("s") ? c.getAttribute("s") : null;
  };
  Sheet.prototype.setStyle = function(ref, s){
    if (s === null || s === undefined) return;
    this.cell(ref, true).setAttribute("s", s);
  };
  Sheet.prototype.hasFormula = function(ref){
    var c = this.cell(ref, false);
    return !!(c && byTag(c, "f").length);
  };
  Sheet.prototype.textOf = function(ref){
    var c = this.cell(ref, false);
    if (!c) return "";
    var t = c.getAttribute("t"), v = byTag(c, "v")[0];
    if (t === "s") return v ? (this.shared[parseInt(v.textContent, 10)] || "") : "";
    if (t === "inlineStr") return byTag(c, "t").map(function(x){ return x.textContent; }).join("");
    return v ? v.textContent : "";
  };
  // Removes the value of a plain cell (never touches formula cells), keeping its style.
  Sheet.prototype.clear = function(ref){
    var c = this.cell(ref, false);
    if (!c || byTag(c, "f").length) return;
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute("t");
  };
  Sheet.prototype.setNumber = function(ref, n){
    var c = this.cell(ref, true);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute("t");
    var v = this.doc.createElementNS(this.ns, "v");
    v.textContent = String(Math.round(n * 1e10) / 1e10);
    c.appendChild(v);
  };
  Sheet.prototype.setText = function(ref, text){
    var c = this.cell(ref, true);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.setAttribute("t", "inlineStr");
    var is = this.doc.createElementNS(this.ns, "is"), t = this.doc.createElementNS(this.ns, "t");
    t.setAttributeNS(NS_XML, "xml:space", "preserve");
    t.textContent = String(text);
    is.appendChild(t);
    c.appendChild(is);
  };
  // Formula results cached in the template are dropped; Excel recalculates on open.
  Sheet.prototype.dropCachedResults = function(){
    byTag(this.data, "c").forEach(function(c){
      if (!byTag(c, "f").length) return;
      byTag(c, "v").forEach(function(v){ c.removeChild(v); });
      if (c.getAttribute("t") !== "array") c.removeAttribute("t");
    });
  };
  Sheet.prototype.merges = function(){
    return byTag(this.doc, "mergeCell").map(function(m){ return m.getAttribute("ref"); });
  };
  Sheet.prototype.removeMerges = function(test){
    byTag(this.doc, "mergeCell").forEach(function(m){ if (test(m.getAttribute("ref"))) m.parentNode.removeChild(m); });
    this.updateMergeCount();
  };
  Sheet.prototype.addMerge = function(ref){
    var holder = byTag(this.doc, "mergeCells")[0];
    if (!holder){
      holder = this.doc.createElementNS(this.ns, "mergeCells");
      this.data.parentNode.insertBefore(holder, this.data.nextSibling);
    }
    var m = this.doc.createElementNS(this.ns, "mergeCell");
    m.setAttribute("ref", ref);
    holder.appendChild(m);
    this.updateMergeCount();
  };
  Sheet.prototype.updateMergeCount = function(){
    var holder = byTag(this.doc, "mergeCells")[0];
    if (holder) holder.setAttribute("count", String(byTag(holder, "mergeCell").length));
  };
  Sheet.prototype.setDimension = function(ref){
    var d = byTag(this.doc, "dimension")[0];
    if (d) d.setAttribute("ref", ref);
  };

  // ---------- Portuguese calendar ----------
  function easter(y){
    var a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4;
    var f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
    var i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
    var month = Math.floor((h + l - 7 * m + 114) / 31), day = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(y, month - 1, day);
  }
  // National holidays plus Carnival (tolerance day usually given), as "M-D".
  function holidays(y){
    var set = {};
    ["1-1","4-25","5-1","6-10","8-15","10-5","11-1","12-1","12-8","12-25"].forEach(function(k){ set[k] = true; });
    var e = easter(y);
    [-47, -2, 0, 60].forEach(function(offset){
      var d = new Date(e.getFullYear(), e.getMonth(), e.getDate() + offset);
      set[(d.getMonth() + 1) + "-" + d.getDate()] = true;
    });
    return set;
  }
  function workdays(y, m){ // m: 1-12
    var hol = holidays(y), days = new Date(y, m, 0).getDate(), n = 0;
    for (var d = 1; d <= days; d++){
      var wd = new Date(y, m - 1, d).getDay();
      if (wd !== 0 && wd !== 6 && !hol[m + "-" + d]) n++;
    }
    return n;
  }
  function isoWeek(date){
    var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    var day = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - day);
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  }
  var WEEKDAY_ABBR = ["Dom","Seg","Ter","Qua","Qui","Sex","Sab"];
  var MONTHS_PLAIN = ["janeiro","fevereiro","marco","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];

  function mode(list){
    var count = {}, best = null;
    list.forEach(function(v){ if (v === null) return; count[v] = (count[v] || 0) + 1; if (best === null || count[v] > count[best]) best = v; });
    return best;
  }

  // ---------- cell comments (notes), written like the ones Excel makes ----------
  var NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
  var NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  var REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
  var REL_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
  var CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
  var CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";

  function euro(n){
    var parts = Math.abs(n).toFixed(2).split(".");
    return (n < 0 ? "-" : "") + parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, " ") + "," + parts[1] + " €";
  }
  function escXml(s){
    return String(s).replace(/[&<>"']/g, function(c){
      return { "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&apos;" }[c];
    });
  }
  // Author shown in bold at the top of each comment: the template's own, else the name in F1.
  function commentAuthor(pkg, firstSheet){
    var names = pkg.entries.map(function(e){ return e.name; }).filter(function(n){ return /^xl\/comments\d+\.xml$/.test(n); });
    var fromSheet = firstSheet.textOf("F1").trim();
    if (!names.length) return Promise.resolve(fromSheet || "Despesas");
    return pkg.text(names[0]).then(function(t){
      var a = byTag(parseXml(t), "author")[0];
      return (a && a.textContent.trim()) || fromSheet || "Despesas";
    });
  }
  function resolveTarget(sheetPath, target){
    if (target.charAt(0) === "/") return target.slice(1);
    var parts = sheetPath.split("/").slice(0, -1);
    target.split("/").forEach(function(seg){
      if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg);
    });
    return parts.join("/");
  }
  function freePartName(pkg, prefix, ext){
    for (var n = 1; ; n++) if (!pkg.has(prefix + n + ext)) return prefix + n + ext;
  }
  function commentsXml(author, comments){
    var runStyle = '<rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>';
    var boldStyle = '<rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>';
    return XML_DECL +
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<authors><author>' + escXml(author) + '</author></authors><commentList>' +
      comments.map(function(c){
        return '<comment ref="' + c.ref + '" authorId="0"><text>' +
          '<r>' + boldStyle + '<t>' + escXml(author) + ':</t></r>' +
          '<r>' + runStyle + '<t xml:space="preserve">\n' + escXml(c.text) + '</t></r>' +
          '</text></comment>';
      }).join("") +
      '</commentList></comments>';
  }
  function vmlXml(sheetNumber, comments){
    var shapes = comments.map(function(c, k){
      var p = splitRef(c.ref), row = p.row - 1, col = colIndex(p.col) - 1;
      return '<v:shape id="_x0000_s' + (sheetNumber * 1024 + k + 1) + '" type="#_x0000_t202" style="position:absolute;' +
        'margin-left:59.25pt;margin-top:1.5pt;width:150pt;height:60pt;z-index:' + (k + 1) + ';visibility:hidden" ' +
        'fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">' +
        '<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/>' +
        '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
        '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
        '<x:Anchor>' + [col + 1, 15, Math.max(0, row - 1), 10, col + 4, 15, row + 4, 4].join(", ") + '</x:Anchor>' +
        '<x:AutoFill>False</x:AutoFill><x:Row>' + row + '</x:Row><x:Column>' + col + '</x:Column></x:ClientData></v:shape>';
    }).join("");
    return '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="' + sheetNumber + '"/></o:shapelayout>' +
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
      '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' + shapes + '</xml>';
  }
  function setOverride(ct, partName, type){
    var found = byTag(ct, "Override").filter(function(o){ return o.getAttribute("PartName") === "/" + partName; })[0];
    if (!type){ if (found) found.parentNode.removeChild(found); return; }
    if (found){ found.setAttribute("ContentType", type); return; }
    var o = ct.createElementNS(NS_CT, "Override");
    o.setAttribute("PartName", "/" + partName);
    o.setAttribute("ContentType", type);
    ct.documentElement.appendChild(o);
  }
  function ensureDefault(ct, ext, type){
    var has = byTag(ct, "Default").some(function(d){ return (d.getAttribute("Extension") || "").toLowerCase() === ext; });
    if (has) return;
    var d = ct.createElementNS(NS_CT, "Default");
    d.setAttribute("Extension", ext);
    d.setAttribute("ContentType", type);
    ct.documentElement.insertBefore(d, ct.documentElement.firstChild);
  }

  // Replaces the sheet's comments with `comments` ([{ ref, text }]); none → the sheet has no comments.
  function writeSheetComments(pkg, ct, sheet, sheetPath, sheetNumber, comments, author){
    var relsPath = sheetPath.replace(/([^\/]+)$/, "_rels/$1.rels");
    return pkg.text(relsPath).then(function(relsText){
      var rels = parseXml(relsText || '<Relationships xmlns="' + NS_PKG_REL + '"/>');
      var list = byTag(rels, "Relationship");
      var commentsRel = list.filter(function(r){ return r.getAttribute("Type") === REL_COMMENTS; })[0];
      var vmlRel = list.filter(function(r){ return r.getAttribute("Type") === REL_VML; })[0];
      var legacy = byTag(sheet.doc, "legacyDrawing")[0];

      if (!comments.length){
        if (commentsRel){
          var cPath = resolveTarget(sheetPath, commentsRel.getAttribute("Target"));
          pkg.remove(cPath); setOverride(ct, cPath, null);
          commentsRel.parentNode.removeChild(commentsRel);
        }
        if (vmlRel){
          pkg.remove(resolveTarget(sheetPath, vmlRel.getAttribute("Target")));
          vmlRel.parentNode.removeChild(vmlRel);
        }
        if (legacy) legacy.parentNode.removeChild(legacy);
        if (relsText) pkg.setText(relsPath, serializeXml(rels));
        return;
      }

      function newRelId(){
        var used = byTag(rels, "Relationship").map(function(r){ return r.getAttribute("Id"); });
        for (var n = 1; ; n++) if (used.indexOf("rId" + n) < 0) return "rId" + n;
      }
      function addRel(type, partPath){
        var r = rels.createElementNS(NS_PKG_REL, "Relationship");
        r.setAttribute("Id", newRelId());
        r.setAttribute("Type", type);
        r.setAttribute("Target", "../" + partPath.replace(/^xl\//, ""));
        rels.documentElement.appendChild(r);
        return r;
      }
      var commentsPath = commentsRel ? resolveTarget(sheetPath, commentsRel.getAttribute("Target")) : freePartName(pkg, "xl/comments", ".xml");
      if (!commentsRel) commentsRel = addRel(REL_COMMENTS, commentsPath);
      var vmlPath = vmlRel ? resolveTarget(sheetPath, vmlRel.getAttribute("Target")) : freePartName(pkg, "xl/drawings/vmlDrawing", ".vml");
      if (!vmlRel) vmlRel = addRel(REL_VML, vmlPath);

      pkg.setText(commentsPath, commentsXml(author, comments));
      pkg.setText(vmlPath, vmlXml(sheetNumber, comments));
      setOverride(ct, commentsPath, CT_COMMENTS);
      ensureDefault(ct, "vml", CT_VML);

      if (!legacy){
        legacy = sheet.doc.createElementNS(sheet.ns, "legacyDrawing");
        // Schema order: … drawing, legacyDrawing, legacyDrawingHF, picture, oleObjects, controls, …, tableParts, extLst
        var drawing = byTag(sheet.doc, "drawing")[0];
        var after = ["legacyDrawingHF","picture","oleObjects","controls","webPublishItems","tableParts","extLst"]
          .map(function(t){ return byTag(sheet.doc, t)[0]; }).filter(Boolean)[0];
        var root = sheet.doc.documentElement;
        if (drawing && drawing.parentNode === root) root.insertBefore(legacy, drawing.nextSibling);
        else if (after) root.insertBefore(legacy, after);
        else root.appendChild(legacy);
      }
      legacy.setAttributeNS(NS_REL, "r:id", vmlRel.getAttribute("Id"));
      pkg.setText(relsPath, serializeXml(rels));
    });
  }

  // ---------- Despesas: monthly map, one row per day ----------
  // items: [{ date:"AAAA-MM-DD", col:"D".."N", amount:n, note?:"text for column K" }]
  var DAY_FIRST_ROW = 5, DAY_LAST_ROW = 35;
  var DAY_COLS = ["B","C","D","E","F","G","H","I","J","K","L","M","N","O"];
  var VALUE_COLS = ["D","E","F","G","H","I","J","K","L","M","N"];

  function fillMonthSheet(sheet, year, monthIdx, items, aStyles){
    var days = new Date(year, monthIdx + 1, 0).getDate();
    // Styles of weekday / weekend rows, as the template uses them.
    var weekday = {}, weekend = {};
    DAY_COLS.forEach(function(c){ weekday[c] = []; weekend[c] = []; });
    for (var r = DAY_FIRST_ROW; r <= DAY_LAST_ROW; r++){
      var w = plain(sheet.textOf("C" + r));
      if (!w) continue;
      var target = (w === "sab" || w === "dom") ? weekend : weekday;
      DAY_COLS.forEach(function(c){ target[c].push(sheet.styleOf(c + r)); });
    }
    DAY_COLS.forEach(function(c){ weekday[c] = mode(weekday[c]); weekend[c] = mode(weekend[c]) || weekday[c]; });

    // Week-number column A: styles for the top / middle / bottom of a merged week, and a one-day week.
    var aMerges = sheet.merges().filter(function(ref){
      var m = /^A(\d+):A(\d+)$/.exec(ref);
      return m && +m[1] >= DAY_FIRST_ROW && +m[2] <= DAY_LAST_ROW;
    });
    aMerges.forEach(function(ref){
      var m = /^A(\d+):A(\d+)$/.exec(ref), a = +m[1], b = +m[2];
      if (!aStyles.top) aStyles.top = sheet.styleOf("A" + a);
      if (!aStyles.bottom) aStyles.bottom = sheet.styleOf("A" + b);
      if (!aStyles.middle && b - a > 1) aStyles.middle = sheet.styleOf("A" + (a + 1));
    });
    sheet.removeMerges(function(ref){ return aMerges.indexOf(ref) >= 0; });

    for (r = DAY_FIRST_ROW; r <= DAY_LAST_ROW; r++){
      var day = r - DAY_FIRST_ROW + 1;
      VALUE_COLS.forEach(function(c){ sheet.clear(c + r); });
      sheet.clear("A" + r);
      if (day > days){ sheet.clear("B" + r); sheet.clear("C" + r); continue; }
      var wd = new Date(year, monthIdx, day).getDay();
      var styles = (wd === 0 || wd === 6) ? weekend : weekday;
      DAY_COLS.forEach(function(c){ if (sheet.cell(c + r, false) || styles[c]) sheet.setStyle(c + r, styles[c]); });
      sheet.setNumber("B" + r, day);
      sheet.setText("C" + r, WEEKDAY_ABBR[wd]);
    }

    // Weeks (Monday to Sunday) with their ISO week number, merged like the template.
    var start = 1;
    while (start <= days){
      var end = start;
      while (end < days && new Date(year, monthIdx, end + 1).getDay() !== 1) end++;
      var top = DAY_FIRST_ROW + start - 1, bottom = DAY_FIRST_ROW + end - 1;
      sheet.setNumber("A" + top, isoWeek(new Date(year, monthIdx, start)));
      if (bottom > top){
        sheet.setStyle("A" + top, aStyles.top);
        for (var x = top + 1; x < bottom; x++) sheet.setStyle("A" + x, aStyles.middle || aStyles.top);
        sheet.setStyle("A" + bottom, aStyles.bottom || aStyles.top);
        sheet.addMerge("A" + top + ":A" + bottom);
      } else {
        sheet.setStyle("A" + top, aStyles.single || aStyles.top);
      }
      start = end + 1;
    }

    // Amounts: same day and column are added together; notes go to column K;
    // descriptions become the cell's comment.
    var sums = {}, notes = {}, texts = {};
    items.forEach(function(it){
      var day = parseInt(it.date.slice(8, 10), 10);
      var key = it.col + (DAY_FIRST_ROW + day - 1);
      sums[key] = (sums[key] || 0) + (Number(it.amount) || 0);
      texts[key] = texts[key] || [];
      texts[key].push({ amount:Number(it.amount) || 0, text:String(it.comment || "").trim() });
      if (it.note){
        var k = "K" + (DAY_FIRST_ROW + day - 1);
        notes[k] = notes[k] || [];
        if (notes[k].indexOf(it.note) < 0) notes[k].push(it.note);
      }
    });
    Object.keys(sums).forEach(function(ref){ sheet.setNumber(ref, Math.round(sums[ref] * 100) / 100); });
    Object.keys(notes).forEach(function(ref){ sheet.setText(ref, notes[ref].join(" / ")); });

    var comments = [];
    Object.keys(texts).forEach(function(ref){
      var list = texts[ref];
      if (!list.some(function(x){ return x.text; })) return;
      var text = list.length === 1 ? list[0].text : list.map(function(x){
        return euro(x.amount) + (x.text ? ": " + x.text : "");
      }).join("\n");
      comments.push({ ref:ref, text:text });
    });

    // "Subsídios de almoço": working days of this month in the formula (e.g. 22-J40-G38).
    byTag(sheet.doc, "c").forEach(function(c){
      var f = byTag(c, "f")[0];
      if (f && /^\s*\+?\d+\s*-\s*J\d+\s*-\s*G\d+\s*$/i.test(f.textContent)){
        f.textContent = f.textContent.replace(/^\s*\+?\d+/, String(workdays(year, monthIdx + 1)));
      }
    });
    return comments;
  }

  function buildDespesas(templateBuffer, year, items){
    var pkg = new Package(templateBuffer);
    return listSheets(pkg).then(function(info){
      var monthSheets = [];
      info.sheets.forEach(function(s){
        var idx = MONTHS_PLAIN.indexOf(plain(s.name));
        if (idx >= 0) monthSheets.push({ idx:idx, path:s.path });
      });
      if (monthSheets.length < 12) throw new Error("o modelo Despesas.xlsx não tem as 12 folhas dos meses");
      monthSheets.sort(function(a, b){ return a.idx - b.idx; });
      return Promise.all(monthSheets.map(function(s){ return pkg.text(s.path); })).then(function(texts){
        var sheets = texts.map(function(t){ return new Sheet(parseXml(t), info.shared); });
        // One-day-week style, taken from any month where it occurs in the template.
        var aStyles = {};
        sheets.forEach(function(sh){
          if (aStyles.single) return;
          var covered = {};
          sh.merges().forEach(function(ref){
            var m = /^A(\d+):A(\d+)$/.exec(ref);
            if (m) for (var r = +m[1]; r <= +m[2]; r++) covered[r] = true;
          });
          for (var r = DAY_FIRST_ROW; r <= DAY_LAST_ROW; r++){
            if (!covered[r] && sh.textOf("A" + r) && sh.textOf("C" + r)){ aStyles.single = sh.styleOf("A" + r); break; }
          }
        });
        return commentAuthor(pkg, sheets[0]).then(function(author){
          return pkg.text("[Content_Types].xml").then(function(ctText){
            var ct = parseXml(ctText);
            var chain = Promise.resolve();
            sheets.forEach(function(sh, i){
              chain = chain.then(function(){
                var monthKey = year + "-" + String(i + 1).padStart(2, "0");
                var monthItems = items.filter(function(it){ return it.date.slice(0, 7) === monthKey; });
                var comments = fillMonthSheet(sh, year, i, monthItems, aStyles);
                if (!sh.hasFormula("O2")) sh.setNumber("O2", year); // the year (other months use =Janeiro!O2)
                sh.dropCachedResults();
                return writeSheetComments(pkg, ct, sh, monthSheets[i].path, i + 1, comments, author).then(function(){
                  pkg.setText(monthSheets[i].path, serializeXml(sh.doc));
                });
              });
            });
            return chain.then(function(){
              pkg.setText("[Content_Types].xml", serializeXml(ct));
              // Ask Excel to recalculate all formulas when the file is opened.
              var calc = byTag(info.workbookDoc, "calcPr")[0];
              if (calc) calc.setAttribute("fullCalcOnLoad", "1");
              pkg.setText("xl/workbook.xml", serializeXml(info.workbookDoc));
              return pkg.build();
            });
          });
        });
      });
    });
  }

  // ---------- Combustíveis: list of fleet-card fuel purchases ----------
  // rows: [{ card, date:"AAAA-MM-DD", time:"HH:MM", station, zone, name, kms, price, product, liters }]
  var FUEL_HEADERS = {
    ncartao:"card", cartao:"card", data:"date", hora:"time", posto:"station", zona:"zone", nome:"name",
    kms:"kms", km:"kms", preco:"price", valor:"price", produto:"product", quantidadel:"liters", quantidade:"liters", litros:"liters"
  };
  function excelDate(iso){
    var p = iso.split("-").map(Number);
    return (Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000;
  }
  function excelTime(hhmm){
    var m = /^(\d{1,2}):(\d{2})/.exec(hhmm || "");
    return m ? (+m[1]) / 24 + (+m[2]) / 1440 : null;
  }

  function buildCombustiveis(templateBuffer, rows){
    var pkg = new Package(templateBuffer);
    return listSheets(pkg).then(function(info){
      var first = info.sheets[0];
      return pkg.text(first.path).then(function(text){
        var sh = new Sheet(parseXml(text), info.shared);
        var rowNums = Object.keys(sh.rows).map(Number).sort(function(a, b){ return a - b; });
        // Header row: the one with "Data" and "Kms"/"Preço".
        var headerRow = null, cols = {};
        rowNums.some(function(r){
          var found = {};
          byTag(sh.rows[r], "c").forEach(function(c){
            var key = FUEL_HEADERS[plain(sh.textOf(c.getAttribute("r")))];
            if (key && !found[key]) found[key] = splitRef(c.getAttribute("r")).col;
          });
          if (found.date && (found.kms || found.price)){ headerRow = r; cols = found; return true; }
          return false;
        });
        if (!headerRow) throw new Error("não encontrei o cabeçalho (Data, Kms, Preço…) no Combustiveis.xlsx");
        // Example rows right below the header ("Ex.:") are kept as in the template.
        var startRow = headerRow + 1;
        while (/^ex/.test(plain(sh.textOf("A" + startRow)))) startRow++;
        var styleRow = startRow > headerRow + 1 ? startRow - 1 : startRow;
        var styles = {};
        Object.keys(cols).forEach(function(k){ styles[k] = sh.styleOf(cols[k] + styleRow) || sh.styleOf(cols[k] + (styleRow + 1)); });
        var lastRow = rowNums[rowNums.length - 1] || startRow;
        var allCols = Object.keys(cols).map(function(k){ return cols[k]; });
        for (var r = startRow; r <= lastRow; r++) allCols.forEach(function(c){ sh.clear(c + r); });

        rows.forEach(function(it, i){
          var r = startRow + i;
          Object.keys(cols).forEach(function(k){ sh.setStyle(cols[k] + r, styles[k]); });
          function num(k, v){ if (cols[k] && v !== null && v !== undefined && v !== "" && !isNaN(Number(v))) sh.setNumber(cols[k] + r, Number(v)); }
          function txt(k, v){ if (cols[k] && v) sh.setText(cols[k] + r, v); }
          txt("card", it.card);
          if (cols.date && it.date) sh.setNumber(cols.date + r, excelDate(it.date));
          num("time", excelTime(it.time));
          txt("station", it.station);
          txt("zone", it.zone);
          txt("name", it.name);
          num("kms", it.kms);
          num("price", it.price);
          txt("product", it.product);
          num("liters", it.liters);
        });
        var maxRow = Math.max(lastRow, startRow + rows.length - 1);
        var dim = byTag(sh.doc, "dimension")[0];
        if (dim){
          var m = /^([A-Z]+\d+):([A-Z]+)\d+$/.exec(dim.getAttribute("ref") || "");
          if (m) sh.setDimension(m[1] + ":" + m[2] + maxRow);
        }
        sh.dropCachedResults();
        pkg.setText(first.path, serializeXml(sh.doc));
        return pkg.build();
      });
    });
  }

  window.QuadroExcel = {
    buildDespesas: buildDespesas,
    buildCombustiveis: buildCombustiveis,
    workdays: workdays,
    isoWeek: isoWeek
  };
})();
