/* =========================================================================
   SENECA NAVEGA — keyword-based guide engine (no AI / no backend)  [engine v5]
   =========================================================================
   Included the same way on all pages. Each page only needs to call
   SenecaNavega.init({...}) with its own configuration.

   Two display modes, chosen automatically per page based on what you pass in:

   1) FLOATER MODE (recommended — used on index.html):
      Pass floaterEl / floaterImgEl / bubbleEl / resolveElement.
      When Seneca finds an answer, it does NOT print it in the chat — instead
      the Seneca image moves and hovers on top of the element the user should
      click, with a speech bubble giving the exact direction.

   2) FALLBACK CHAT MODE (used automatically if no floater config is given —
      this is what Baseroom/KPIs currently use):
      The answer is printed as a chat bubble with a "Take me there" button
      that navigates/highlights the target using the page's `adapter`.
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

  // ---------- typewriter (letter-by-letter reveal, HTML-safe) ----------
  // Builds the real DOM structure up front (so tags like <strong> are proper,
  // stable elements) and then reveals text-node characters one at a time.
  // This avoids the classic bug of repeatedly reassigning innerHTML, which
  // makes the browser auto-close "unfinished" tags on every step.
  function typewriter(targetEl, html, speed, onDone) {
    var tmp = document.createElement("div");
    tmp.innerHTML = html;
    targetEl.innerHTML = "";

    var queue = [];
    function cloneAndQueue(srcNode, destParent) {
      for (var i = 0; i < srcNode.childNodes.length; i++) {
        var child = srcNode.childNodes[i];
        if (child.nodeType === 3) {
          var newText = document.createTextNode("");
          destParent.appendChild(newText);
          queue.push({ node: newText, full: child.nodeValue });
        } else if (child.nodeType === 1) {
          var newEl = document.createElement(child.tagName);
          for (var a = 0; a < child.attributes.length; a++) {
            newEl.setAttribute(child.attributes[a].name, child.attributes[a].value);
          }
          destParent.appendChild(newEl);
          cloneAndQueue(child, newEl);
        }
      }
    }
    cloneAndQueue(tmp, targetEl);

    var qi = 0, ci = 0;
    function step() {
      if (qi >= queue.length) { if (onDone) onDone(); return; }
      var item = queue[qi];
      var full = item.full;
      if (ci >= full.length) { qi++; ci = 0; step(); return; }
      var code = full.codePointAt(ci);
      var ch = String.fromCodePoint(code);
      item.node.data += ch;
      ci += ch.length;
      var log = targetEl.closest ? targetEl.closest(".seneca-chat__log") : null;
      if (log) log.scrollTop = log.scrollHeight;
      setTimeout(step, speed || 16);
    }
    step();
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

  function addTypedMsg(log, html, speed) {
    var div = el("div", "seneca-msg seneca-msg--bot");
    log.appendChild(div);
    log.scrollTop = log.scrollHeight;
    return new Promise(function (resolve) {
      typewriter(div, html, speed, function () { resolve(div); });
    });
  }

  function crossfadeImage(imgEl, newSrc) {
    if (!imgEl || !newSrc) return;
    imgEl.style.transition = "opacity 0.3s ease";
    imgEl.style.opacity = "0";
    setTimeout(function () {
      imgEl.src = newSrc;
      imgEl.style.opacity = "1";
    }, 260);
  }

  // Swaps the panel's inner content only (crossfade). The panel box itself
  // (background / border / shadow) is never hidden or removed, so it never
  // "appears from zero" — it's always the same box, only what's inside changes.
  function swapPanelContent(panel, buildFn, afterInsert) {
    var old = panel.firstElementChild;
    function insertNew() {
      panel.innerHTML = "";
      var built = buildFn();
      built.style.opacity = "0";
      built.style.transition = "opacity .25s ease";
      panel.appendChild(built);
      void built.offsetWidth;
      requestAnimationFrame(function () { built.style.opacity = "1"; });
      if (afterInsert) afterInsert();
    }
    if (old) {
      old.style.transition = "opacity .15s ease";
      old.style.opacity = "0";
      setTimeout(insertNew, 150);
    } else {
      insertNew();
    }
  }

  // ---------- small talk (greetings / help / thanks — not a location search) ----------
  var SMALL_TALK = [
    { test: /\b(hi|hello|hey|hola|good morning|good afternoon|good evening)\b/,
      reply: "Hello! Tell me what document, folder or dashboard you're looking for and I'll show you exactly where to click." },
    { test: /\b(help|ayuda|how does this work|what can you do|how do you work)\b/,
      reply: "I'm <strong>Séneca Navega</strong>. Ask me about any topic (e.g. \"annual reports\", \"faculty questionnaire\") and I'll point right at where to click to find it." },
    { test: /\b(thanks|thank you|gracias|thx)\b/,
      reply: "You're welcome! Let me know if you need anything else." }
  ];
  function matchSmallTalk(q) {
    var norm = normalize(q);
    for (var i = 0; i < SMALL_TALK.length; i++) {
      if (SMALL_TALK[i].test.test(norm)) return SMALL_TALK[i].reply;
    }
    return null;
  }

  // ---------- init ----------
  function init(config) {
    var page = config.page;                 // "index" | "baseroom" | "kpis"
    var triggerBtn = config.triggerBtn;      // button that activates Seneca
    var panel = config.panelToReplace;       // container that becomes the chat (size never changes)
    var imageEl = config.imageEl;            // idle <img> shown before activation
    var activeImageSrc = config.activeImageSrc;
    var idleMessage = config.idleMessage || "Hi, I'm <strong>Séneca Navega</strong>. Ask me where to find something.";
    var placeholder = config.placeholder || "Ex: where is the risk register...";
    var adapter = config.adapter || {};
    var index = buildIndex(config.index || []);

    // floater mode (optional — index.html uses this)
    var floaterEl = config.floaterEl || null;
    var floaterImgEl = config.floaterImgEl || null;
    var bubbleEl = config.bubbleEl || null;
    var resolveElement = config.resolveElement || null;
    var floaterSize = config.floaterSize || 190;
    var floaterHomeSize = config.floaterHomeSize || Math.round(floaterSize * 0.6);

    var triggerBtnId = triggerBtn.id;
    var originalPanelHTML = panel.innerHTML;
    // Capture the panel's original height (before activation) so the chat can
    // reproduce that exact same outer size — never taller, never shorter.
    var originalPanelHeight = panel.getBoundingClientRect().height;
    var panelComputed = window.getComputedStyle ? window.getComputedStyle(panel) : null;
    var panelPaddingV = panelComputed ? (parseFloat(panelComputed.paddingTop) || 0) + (parseFloat(panelComputed.paddingBottom) || 0) : 0;
    var chatContentHeight = originalPanelHeight ? Math.max(originalPanelHeight - panelPaddingV, 0) : 0;
    var idleImageSrc = imageEl ? imageEl.src : null;

    var active = false;
    var chatRoot = null;
    var logEl = null;

    function wireTrigger() {
      var btn = document.getElementById(triggerBtnId);
      if (btn) btn.addEventListener("click", function () { activate(false); });
    }

    // ---- floater animation helpers ----
    // Grows Seneca from the idle logo's spot, anchored so it expands to the
    // LEFT (away from the panel) instead of over it — it never covers the text.
    function growFloaterFromIdleLogo() {
      if (!floaterEl || !imageEl) return;
      var homeRect = imageEl.getBoundingClientRect();
      var startW = homeRect.width || floaterHomeSize;
      imageEl.style.transition = "opacity 0.35s ease";
      imageEl.style.opacity = "0";
      if (floaterImgEl && activeImageSrc) floaterImgEl.src = activeImageSrc;
      floaterEl.style.transition = "none";
      floaterEl.style.top = homeRect.top + "px";
      floaterEl.style.left = homeRect.left + "px";
      floaterEl.style.width = startW + "px";
      floaterEl.style.opacity = "0";
      floaterEl.style.transform = "scale(0.85)";
      floaterEl.style.display = "block";
      void floaterEl.offsetWidth; // force reflow before animating
      requestAnimationFrame(function () {
        floaterEl.style.transition = "opacity .5s ease, transform .5s ease, width .5s ease, top .5s ease, left .5s ease";
        floaterEl.style.opacity = "1";
        floaterEl.style.transform = "scale(1)";
        var deltaW = floaterHomeSize - startW;
        floaterEl.style.width = floaterHomeSize + "px";
        floaterEl.style.left = (homeRect.left - deltaW) + "px"; // grow leftwards, away from the panel
      });
    }

    function shrinkFloaterToIdleLogo() {
      if (!floaterEl || !imageEl) return;
      hideBubble();
      var homeRect = imageEl.getBoundingClientRect();
      floaterEl.style.transition = "opacity .35s ease, transform .35s ease, width .35s ease, top .45s ease, left .45s ease";
      floaterEl.style.top = homeRect.top + "px";
      floaterEl.style.left = homeRect.left + "px";
      floaterEl.style.width = (homeRect.width || floaterHomeSize) + "px";
      floaterEl.style.opacity = "0";
      floaterEl.style.transform = "scale(0.85)";
      setTimeout(function () {
        floaterEl.style.display = "none";
        imageEl.style.opacity = "1";
      }, 380);
    }

    // Moves Seneca back to its starting spot (small, to the left of the panel)
    // WITHOUT hiding it (still active) — used for small talk so it never stays
    // parked on top of a card/text.
    function returnFloaterHome() {
      if (!floaterEl || !imageEl) return;
      hideBubble();
      var homeRect = imageEl.getBoundingClientRect();
      var currentW = floaterEl.getBoundingClientRect().width || floaterHomeSize;
      var deltaW = floaterHomeSize - (homeRect.width || floaterHomeSize);
      floaterEl.style.transition = "top .45s ease, left .45s ease, width .45s ease";
      floaterEl.style.top = homeRect.top + "px";
      floaterEl.style.left = (homeRect.left - deltaW) + "px";
      floaterEl.style.width = floaterHomeSize + "px";
    }

    function showBubble(html) {
      if (!bubbleEl) return;
      bubbleEl.style.display = "block";
      bubbleEl.style.opacity = "0";
      void bubbleEl.offsetWidth;
      bubbleEl.style.transition = "opacity .3s ease";
      bubbleEl.style.opacity = "1";
      typewriter(bubbleEl, html, 14);
    }

    function hideBubble() {
      if (!bubbleEl) return;
      bubbleEl.style.opacity = "0";
      setTimeout(function () { bubbleEl.style.display = "none"; bubbleEl.innerHTML = ""; }, 250);
    }

    function pointFloaterAt(targetEl, html) {
      if (!floaterEl || !targetEl) return;
      if (targetEl.scrollIntoView) targetEl.scrollIntoView({ behavior: "smooth", block: "center" });
      setTimeout(function () {
        var rect = targetEl.getBoundingClientRect();
        var fw = floaterSize;
        var top = rect.top - fw * 0.55;
        var left = rect.left + Math.min(24, rect.width * 0.1);
        if (top < 8) top = rect.bottom + 10; // not enough room above -> place below instead
        floaterEl.style.transition = "opacity .3s ease, transform .3s ease, width .5s ease, top .5s ease, left .5s ease";
        floaterEl.style.width = fw + "px";
        floaterEl.style.top = top + "px";
        floaterEl.style.left = left + "px";
        showBubble(html);
        targetEl.classList.add("seneca-highlight");
        setTimeout(function () { targetEl.classList.remove("seneca-highlight"); }, 2600);
      }, 380);
    }

    // ---- open / close (the white box itself is always visible; only its
    //      inner content crossfades — width is fixed by CSS, height follows content) ----
    function activate(silent) {
      if (active) return;
      active = true;

      if (floaterEl) growFloaterFromIdleLogo();
      else if (imageEl && activeImageSrc) crossfadeImage(imageEl, activeImageSrc);

      swapPanelContent(panel, function () {
        chatRoot = el("div", "seneca-chat");
        if (chatContentHeight) chatRoot.style.height = chatContentHeight + "px";
        chatRoot.innerHTML =
          '<div class="seneca-chat__header">' +
          '  <button type="button" class="seneca-close-btn" id="senecaCloseBtn" aria-label="Close Séneca Navega">&times;</button>' +
          "</div>" +
          '<div class="seneca-chat__log" id="senecaLog"></div>' +
          '<form class="seneca-chat__form" id="senecaForm" autocomplete="off">' +
          '  <input type="text" id="senecaInput" placeholder="' + placeholder + '">' +
          '  <button type="submit" aria-label="Send">' + ICON_SEND + "</button>" +
          "</form>";
        return chatRoot;
      }, function () {
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

        if (!silent) addTypedMsg(logEl, idleMessage, 16).then(function () { input.focus(); });
        else input.focus();
      });
    }

    function deactivate() {
      active = false;
      if (floaterEl) shrinkFloaterToIdleLogo();
      else if (imageEl && idleImageSrc) crossfadeImage(imageEl, idleImageSrc);

      swapPanelContent(panel, function () {
        var wrap = el("div");
        wrap.innerHTML = originalPanelHTML;
        return wrap;
      }, function () {
        wireTrigger();
      });
    }

    // ---- answering ----
    function makeGotoBtn(target, label) {
      var btn = el("button", "seneca-goto-btn", label || "Take me there →");
      btn.type = "button";
      btn.addEventListener("click", function () { go(target); });
      return btn;
    }

    function renderChoices(candidates) {
      var opts = el("div", "seneca-alts");
      candidates.forEach(function (r) {
        var b = el("button", "seneca-alt-btn", r.entry.label);
        b.type = "button";
        b.addEventListener("click", function () { answerWith(r.entry); });
        opts.appendChild(b);
      });
      logEl.appendChild(opts);
      logEl.scrollTop = logEl.scrollHeight;
    }

    function answerWith(entry) {
      var targetEl = resolveElement ? resolveElement(entry.target) : null;
      if (floaterEl && targetEl) {
        if (page === "index") {
          try { sessionStorage.setItem("senecaPending", JSON.stringify(entry.target)); } catch (e) {}
        }
        pointFloaterAt(targetEl, entry.reply);
      } else {
        addTypedMsg(logEl, entry.reply, 16).then(function (msg) {
          if (entry.target) msg.appendChild(makeGotoBtn(entry.target));
        });
      }
    }

    function handleQuery(q) {
      var smallTalk = matchSmallTalk(q);
      if (smallTalk) {
        addTypedMsg(logEl, smallTalk, 16);
        if (floaterEl) returnFloaterHome();
        return;
      }

      var results = search(q, index, 5);
      if (!results.length) {
        addTypedMsg(
          logEl,
          "I couldn't find anything with those words. Try the name of the document, dashboard or topic (e.g. \"annual reports\", \"faculty questionnaire\").",
          16
        );
        return;
      }

      var top = results[0].score;
      // "ambiguous" = there is more than one candidate close to the top score
      var candidates = results.filter(function (r) { return r.score >= top * 0.75; });

      if (candidates.length > 1) {
        // Only in this case does the chat keep talking: it asks which option the user meant.
        addTypedMsg(logEl, "I found a few things that could match — which one did you mean?", 16)
          .then(function () { renderChoices(candidates.slice(0, 4)); });
        return;
      }

      // Clear match: do NOT answer in the chat — just point at it.
      answerWith(results[0].entry);
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

    // If the user already asked something on the index page, resume here (fallback-mode pages)
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
              addTypedMsg(logEl, "Here's what you were looking for, highlighted on the left.", 16);
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

  window.SenecaNavega = { init: init, search: search, buildIndex: buildIndex, highlightEl: highlightEl, typewriter: typewriter };
})();
