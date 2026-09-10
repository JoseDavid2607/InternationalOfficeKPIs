/* =========================================================================
   SENECA NAVEGA — keyword-based guide engine (no AI / no backend)
   =========================================================================
   Included the same way on all 3 pages (index, Baseroom, KPIs). Each page
   only needs to call SenecaNavega.init({...}) with its own configuration.
   See the "SENECA NAVEGA — integration" blocks inside each HTML file.
   ========================================================================= */
(function () {
  "use strict";

  // ---------- text utilities ----------
  function normalize(s) {
    return (s || "")
      .toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // strip accents
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

  // ---------- UI helpers ----------
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

  // shows 3 bouncing dots, then swaps them for the real message (typing effect)
  function addTypedMsg(log, html, delay) {
    var typing = el("div", "seneca-msg seneca-msg--bot seneca-typing", "<span></span><span></span><span></span>");
    log.appendChild(typing);
    log.scrollTop = log.scrollHeight;
    return new Promise(function (resolve) {
      setTimeout(function () {
        typing.remove();
        resolve(addMsg(log, html, "bot"));
      }, delay || 650);
    });
  }

  function crossfadeImage(imgEl, newSrc) {
    if (!imgEl || !newSrc) return;
    imgEl.classList.add("seneca-img-fade");
    imgEl.style.opacity = "0";
    setTimeout(function () {
      imgEl.src = newSrc;
      imgEl.style.opacity = "1";
    }, 220);
  }

  // ---------- init ----------
  function init(config) {
    var page = config.page;                 // "index" | "baseroom" | "kpis"
    var triggerBtn = config.triggerBtn;      // button that activates Seneca
    var panel = config.panelToReplace;       // container that becomes the chat
    var imageEl = config.imageEl;            // <img> of Seneca to swap
    var activeImageSrc = config.activeImageSrc;
    var idleMessage = config.idleMessage || "Hi, I'm Séneca Navega. Ask me where to find something.";
    var placeholder = config.placeholder || "Ex: where is the risk register...";
    var adapter = config.adapter || {};
    var index = buildIndex(config.index || []);

    var triggerBtnId = triggerBtn.id;
    var originalPanelHTML = panel.innerHTML;      // snapshot so we can restore it on close
    var idleImageSrc = imageEl ? imageEl.src : null;

    var active = false;
    var chatRoot = null;
    var logEl = null;

    function wireTrigger() {
      var btn = document.getElementById(triggerBtnId);
      if (btn) btn.addEventListener("click", function () { activate(false); });
    }

    function activate(silent) {
      if (active) return;
      active = true;
      if (imageEl && activeImageSrc) crossfadeImage(imageEl, activeImageSrc);

      panel.classList.add("seneca-panel-fade");
      panel.style.opacity = "0";
      setTimeout(function () {
        panel.innerHTML = "";
        chatRoot = el("div", "seneca-chat");
        chatRoot.innerHTML =
          '<div class="seneca-chat__header">' +
          '  <span class="seneca-chat__title">🧭 Séneca Navega</span>' +
          '  <button type="button" class="seneca-close-btn" id="senecaCloseBtn" aria-label="Close Séneca Navega">&times;</button>' +
          "</div>" +
          '<div class="seneca-chat__log" id="senecaLog"></div>' +
          '<form class="seneca-chat__form" id="senecaForm" autocomplete="off">' +
          '  <input type="text" id="senecaInput" placeholder="' + placeholder + '">' +
          '  <button type="submit" aria-label="Send">' + ICON_SEND + "</button>" +
          "</form>";
        panel.appendChild(chatRoot);
        panel.style.opacity = "1";

        logEl = chatRoot.querySelector("#senecaLog");
        chatRoot.querySelector("#senecaCloseBtn").addEventListener("click", deactivate);

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

        if (!silent) {
          addTypedMsg(logEl, idleMessage, 500).then(function () { input.focus(); });
        } else {
          input.focus();
        }
      }, 180);
    }

    function deactivate() {
      active = false;
      if (imageEl && idleImageSrc) crossfadeImage(imageEl, idleImageSrc);
      panel.style.opacity = "0";
      setTimeout(function () {
        panel.innerHTML = originalPanelHTML;
        panel.style.opacity = "1";
        wireTrigger();
      }, 180);
    }

    function makeGotoBtn(target, label) {
      var btn = el("button", "seneca-goto-btn", label || "Take me there →");
      btn.type = "button";
      btn.addEventListener("click", function () { go(target); });
      return btn;
    }

    function handleQuery(q) {
      var results = search(q, index, 3);
      if (!results.length) {
        addTypedMsg(
          logEl,
          "I couldn't find anything with those words. Try the name of the document, dashboard or topic (e.g. \"annual reports\", \"faculty questionnaire\").",
          650
        );
        return;
      }
      var best = results[0].entry;
      addTypedMsg(logEl, best.reply, 650).then(function (msg) {
        if (best.target) msg.appendChild(makeGotoBtn(best.target));

        if (results.length > 1) {
          var alts = el("div", "seneca-alts", "<span>Did you mean:</span>");
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
      });
    }

    function go(target) {
      if (page === "index") {
        try { sessionStorage.setItem("senecaPending", JSON.stringify(target)); } catch (e) {}
        location.href = target.page === "baseroom" ? config.baseroomUrl : config.kpisUrl;
        return;
      }
      if (adapter.goTo) adapter.goTo(target);
    }

    wireTrigger();

    // If the user already asked something on the index page, resume the conversation here
    if (page !== "index") {
      var pendingRaw = null;
      try { pendingRaw = sessionStorage.getItem("senecaPending"); } catch (e) {}
      if (pendingRaw) {
        try { sessionStorage.removeItem("senecaPending"); } catch (e) {}
        try {
          var target = JSON.parse(pendingRaw);
          if (target.page === page) {
            activate(true);
            setTimeout(function () {
              addTypedMsg(logEl, "Here's what you were looking for — highlighted on the left. 👇", 500);
              if (adapter.goTo) adapter.goTo(target);
            }, 220);
          }
        } catch (e) {}
      }
    }

    return { activate: activate, deactivate: deactivate };
  }

  function highlightEl(elm) {
    if (!elm) return;
    elm.scrollIntoView({ behavior: "smooth", block: "center" });
    elm.classList.add("seneca-highlight");
    setTimeout(function () { elm.classList.remove("seneca-highlight"); }, 2600);
  }

  window.SenecaNavega = { init: init, search: search, buildIndex: buildIndex, highlightEl: highlightEl };
})();
