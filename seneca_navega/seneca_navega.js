/* =========================================================================
   SENECA NAVEGA — motor de guía por palabras clave (sin IA / sin backend)
   =========================================================================
   Se incluye igual en las 3 paginas (index, Baseroom, KPIs). Cada pagina
   solo debe llamar a SenecaNavega.init({...}) con su propia configuracion.
   Ver los bloques "SENECA NAVEGA — integracion" dentro de cada HTML.
   ========================================================================= */
(function () {
  "use strict";

  // ---------- utilidades de texto ----------
  function normalize(s) {
    return (s || "")
      .toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // quita tildes
      .replace(/[^a-z0-9\s]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  var STOPWORDS = new Set(
    ("de la el los las un una unos unas y o a en para por con sin sobre " +
     "donde cual como que es son the a an of for in on to and or where " +
     "is are how what which do does i need necesito quiero busco encuentro " +
     "encontrar puedo hay tiene tengo mi me").split(" ")
  );

  function tokenize(s) {
    return normalize(s)
      .split(" ")
      .filter(function (w) { return w && w.length > 1 && !STOPWORDS.has(w); });
  }

  function buildIndex(rawEntries) {
    return rawEntries.map(function (e) {
      var text = [e.label, e.keywords || ""].join(" ");
      return Object.assign({}, e, { _tokens: tokenize(text) });
    });
  }

  function scoreEntry(queryTokens, entry) {
    var score = 0;
    queryTokens.forEach(function (qt) {
      entry._tokens.forEach(function (et) {
        if (et === qt) score += 2;
        else if (et.length > 3 && qt.length > 3 && (et.indexOf(qt) !== -1 || qt.indexOf(et) !== -1)) score += 1;
      });
    });
    return score;
  }

  function search(query, index, limit) {
    var qTokens = tokenize(query);
    if (!qTokens.length) return [];
    var scored = index
      .map(function (e) { return { entry: e, score: scoreEntry(qTokens, e) }; })
      .filter(function (r) { return r.score > 0; })
      .sort(function (a, b) { return b.score - a.score; });
    return scored.slice(0, limit || 3);
  }

  // ---------- UI ----------
  var ICON_SEND =
    '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" ' +
    'stroke-linecap="round" stroke-linejoin="round"><path d="M22 2 11 13"/><path d="M22 2 15 22l-4-9-9-4 20-7Z"/></svg>';

  function el(tag, cls, html) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (html !== undefined) n.innerHTML = html;
    return n;
  }

  function addMsg(log, html, who) {
    var div = el("div", "seneca-msg seneca-msg--" + (who || "bot"), html);
    log.appendChild(div);
    log.scrollTop = log.scrollHeight;
    return div;
  }

  // ---------- init ----------
  function init(config) {
    var page = config.page;                 // "index" | "baseroom" | "kpis"
    var triggerBtn = config.triggerBtn;      // boton que activa a Seneca
    var panel = config.panelToReplace;       // contenedor donde aparece el chat
    var imageEl = config.imageEl;            // <img> de Seneca a cambiar
    var activeImageSrc = config.activeImageSrc;
    var idleMessage = config.idleMessage || "Hola, soy Seneca Navega. Pregúntame donde encontrar algo.";
    var placeholder = config.placeholder || "Ej: donde esta el risk register...";
    var adapter = config.adapter || {};
    var index = buildIndex(config.index || []);

    var active = false;
    var chatRoot = null;
    var logEl = null;

    function activate(silent) {
      if (active) return;
      active = true;
      if (imageEl && activeImageSrc) imageEl.src = activeImageSrc;
      panel.innerHTML = "";
      chatRoot = el("div", "seneca-chat");
      chatRoot.innerHTML =
        '<div class="seneca-chat__log" id="senecaLog"></div>' +
        '<form class="seneca-chat__form" id="senecaForm" autocomplete="off">' +
        '  <input type="text" id="senecaInput" placeholder="' + placeholder + '">' +
        '  <button type="submit" aria-label="Enviar">' + ICON_SEND + "</button>" +
        "</form>";
      panel.appendChild(chatRoot);
      logEl = chatRoot.querySelector("#senecaLog");
      if (!silent) addMsg(logEl, idleMessage, "bot");
      var form = chatRoot.querySelector("#senecaForm");
      var input = chatRoot.querySelector("#senecaInput");
      form.addEventListener("submit", function (e) {
        e.preventDefault();
        var q = input.value.trim();
        if (!q) return;
        addMsg(logEl, q, "user");
        input.value = "";
        handleQuery(q);
      });
    }

    function makeGotoBtn(target, label) {
      var btn = el("button", "seneca-goto-btn", label || "Llévame ahí →");
      btn.type = "button";
      btn.addEventListener("click", function () { go(target); });
      return btn;
    }

    function handleQuery(q) {
      var results = search(q, index, 3);
      if (!results.length) {
        addMsg(
          logEl,
          "No encontré nada con esas palabras. Prueba con el nombre del documento, dashboard o tema (ej: \"reportes anuales\", \"encuesta de facultad\").",
          "bot"
        );
        return;
      }
      var best = results[0].entry;
      var msg = addMsg(logEl, best.reply, "bot");
      if (best.target) msg.appendChild(makeGotoBtn(best.target));

      if (results.length > 1) {
        var alts = el("div", "seneca-alts", "<span>¿O tal vez buscabas?</span>");
        results.slice(1).forEach(function (r) {
          var b = el("button", "seneca-alt-btn", r.entry.label);
          b.type = "button";
          b.addEventListener("click", function () {
            var m = addMsg(logEl, r.entry.reply, "bot");
            if (r.entry.target) m.appendChild(makeGotoBtn(r.entry.target));
          });
          alts.appendChild(b);
        });
        logEl.appendChild(alts);
        logEl.scrollTop = logEl.scrollHeight;
      }
    }

    function go(target) {
      if (page === "index") {
        try { sessionStorage.setItem("senecaPending", JSON.stringify(target)); } catch (e) {}
        location.href = target.page === "baseroom" ? config.baseroomUrl : config.kpisUrl;
        return;
      }
      if (adapter.goTo) adapter.goTo(target);
    }

    triggerBtn.addEventListener("click", function () { activate(false); });

    // Si venimos de una pregunta hecha en el index, retomamos la conversacion aqui
    if (page !== "index") {
      var pendingRaw = null;
      try { pendingRaw = sessionStorage.getItem("senecaPending"); } catch (e) {}
      if (pendingRaw) {
        try { sessionStorage.removeItem("senecaPending"); } catch (e) {}
        try {
          var target = JSON.parse(pendingRaw);
          if (target.page === page) {
            activate(true);
            addMsg(logEl, "Aquí tienes lo que buscabas, resaltado a la izquierda. 👇", "bot");
            if (adapter.goTo) adapter.goTo(target);
          }
        } catch (e) {}
      }
    }

    return { activate: activate };
  }

  function highlightEl(elm) {
    if (!elm) return;
    elm.scrollIntoView({ behavior: "smooth", block: "center" });
    elm.classList.add("seneca-highlight");
    setTimeout(function () { elm.classList.remove("seneca-highlight"); }, 2600);
  }

  window.SenecaNavega = { init: init, search: search, buildIndex: buildIndex, highlightEl: highlightEl };
})();
