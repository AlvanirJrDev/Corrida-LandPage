/**
 * Vitrine em destaque: logos grandes alternando automaticamente (data-spotlight-interval em ms).
 */
(function () {
    "use strict";
  
    var DEFAULT_MS = 4500;
    var FADE_MS = 300;
  
    function $(id) {
      return document.getElementById(id);
    }
  
    function prefersReducedMotion() {
      return window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    }
  
    function escapeAttr(s) {
      return String(s)
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;");
    }
  
    function init() {
      var root = $("sponsors-spotlight");
      var list = $("sponsors-showcase-list");
      var panel = $("sponsors-spotlight-panel");
      var captionEl = $("sponsors-spotlight-caption");
      var dotsEl = $("sponsors-spotlight-dots");
      var toggleBtn = $("sponsors-spotlight-toggle");
      var announcer = $("sponsors-spotlight-announcer");
      if (!root || !list || !panel || !captionEl || !dotsEl || !toggleBtn || !announcer) return;
  
      var tiles = list.querySelectorAll(".sponsor-tile");
      if (!tiles.length) return;
  
      var items = [];
      for (var t = 0; t < tiles.length; t++) {
        var tile = tiles[t];
        var img = tile.querySelector(".sponsor-tile__img");
        var nameEl = tile.querySelector(".sponsor-tile__name");
        if (!img || !nameEl) continue;
        items.push({
          src: img.getAttribute("src") || "",
          alt: img.getAttribute("alt") || "",
          name: (nameEl.textContent || "").trim(),
          imgClass: img.className || "",
        });
      }
      if (!items.length) return;
  
      var n = items.length;
      var idx = 0;
      var timer = null;
      var paused = false;
      var reduced = prefersReducedMotion();
  
      var intervalMs = parseInt(root.getAttribute("data-spotlight-interval") || "", 10);
      if (isNaN(intervalMs) || intervalMs < 3000) intervalMs = DEFAULT_MS;
  
      var prevBtn = root.querySelector(".sponsors-spotlight__arrow--prev");
      var nextBtn = root.querySelector(".sponsors-spotlight__arrow--next");
      var pauseLab = toggleBtn.querySelector(".sponsors-spotlight__toggle-label--pause");
      var playLab = toggleBtn.querySelector(".sponsors-spotlight__toggle-label--play");
  
      function announce(msg) {
        announcer.textContent = "";
        announcer.textContent = msg;
      }
  
      function syncToggle() {
        if (!pauseLab || !playLab) return;
        var showPlay = paused || reduced;
        pauseLab.hidden = showPlay;
        playLab.hidden = !showPlay;
        toggleBtn.setAttribute("aria-pressed", showPlay ? "true" : "false");
      }
  
      function clearTimer() {
        if (timer) {
          clearInterval(timer);
          timer = null;
        }
      }
  
      function scheduleTimer() {
        clearTimer();
        if (reduced || paused || n <= 1) return;
        timer = window.setInterval(function () {
          goTo((idx + 1) % n, false);
        }, intervalMs);
      }
  
      function buildDots() {
        dotsEl.innerHTML = "";
        for (var d = 0; d < n; d++) {
          (function (j) {
            var b = document.createElement("button");
            b.type = "button";
            b.className = "sponsors-spotlight__dot";
            b.setAttribute("role", "tab");
            b.setAttribute("aria-selected", j === 0 ? "true" : "false");
            b.setAttribute("aria-label", items[j].name + " — " + (j + 1) + " de " + n);
            b.addEventListener("click", function () {
              goTo(j, true);
            });
            dotsEl.appendChild(b);
          })(d);
        }
      }
  
      function updateDots() {
        var dots = dotsEl.querySelectorAll(".sponsors-spotlight__dot");
        for (var i = 0; i < dots.length; i++) {
          var on = i === idx;
          dots[i].setAttribute("aria-selected", on ? "true" : "false");
          dots[i].classList.toggle("sponsors-spotlight__dot--active", on);
        }
      }
  
      function renderInner(i) {
        var data = items[i];
        var imgClass = "sponsors-spotlight__img";
        if (data.imgClass.indexOf("sponsor-tile__img--mundo-pet") !== -1) {
          imgClass += " sponsors-spotlight__img--mundo-pet";
        }
        if (data.imgClass.indexOf("sponsor-tile__img--gean-veiculos") !== -1) {
          imgClass += " sponsors-spotlight__img--gean-veiculos";
        }
        if (data.imgClass.indexOf("sponsor-tile__img--marcelo-personal") !== -1) {
          imgClass += " sponsors-spotlight__img--marcelo-personal";
        }
        var figClass = "sponsors-spotlight__figure";
        if (data.imgClass.indexOf("sponsor-tile__img--marcelo-personal") !== -1) {
          figClass += " sponsors-spotlight__figure--dark";
        }
        panel.innerHTML =
          '<div class="' +
          figClass +
          '">' +
          '<img class="' +
          imgClass +
          '" src="' +
          escapeAttr(data.src) +
          '" alt="' +
          escapeAttr(data.alt) +
          '" width="480" height="280" decoding="async" />' +
          "</div>";
        captionEl.textContent = data.name;
        updateDots();
        announce("Destaque: " + data.name + ". " + (i + 1) + " de " + n + ".");
      }
  
      function goTo(nextIdx, userClick) {
        var next = ((nextIdx % n) + n) % n;
        if (next === idx && panel.querySelector(".sponsors-spotlight__img")) return;
  
        function applyIndex() {
          idx = next;
          renderInner(idx);
          if (userClick) {
            clearTimer();
            scheduleTimer();
          }
        }
  
        if (reduced) {
          applyIndex();
          return;
        }
  
        panel.classList.add("sponsors-spotlight__panel--out");
        window.setTimeout(function () {
          applyIndex();
          panel.classList.remove("sponsors-spotlight__panel--out");
          void panel.offsetWidth;
          panel.classList.add("sponsors-spotlight__panel--in");
          window.setTimeout(function () {
            panel.classList.remove("sponsors-spotlight__panel--in");
          }, FADE_MS);
        }, FADE_MS);
      }
  
      root.removeAttribute("hidden");
      root.classList.add("sponsors-spotlight--ready");
      buildDots();
      renderInner(0);
      if (!reduced) {
        panel.classList.add("sponsors-spotlight__panel--in");
        window.setTimeout(function () {
          panel.classList.remove("sponsors-spotlight__panel--in");
        }, FADE_MS);
      }
  
      if (n <= 1) {
        if (prevBtn) prevBtn.disabled = true;
        if (nextBtn) nextBtn.disabled = true;
        toggleBtn.hidden = true;
        dotsEl.hidden = true;
      } else {
        if (prevBtn)
          prevBtn.addEventListener("click", function () {
            goTo(idx - 1, true);
          });
        if (nextBtn)
          nextBtn.addEventListener("click", function () {
            goTo(idx + 1, true);
          });
      }
  
      syncToggle();
      if (!reduced) {
        toggleBtn.addEventListener("click", function () {
          paused = !paused;
          syncToggle();
          if (paused) clearTimer();
          else scheduleTimer();
        });
        scheduleTimer();
      }
  
      panel.addEventListener("keydown", function (e) {
        if (n <= 1) return;
        if (e.key === "ArrowLeft") {
          e.preventDefault();
          goTo(idx - 1, true);
        } else if (e.key === "ArrowRight") {
          e.preventDefault();
          goTo(idx + 1, true);
        } else if (e.key === "Home") {
          e.preventDefault();
          goTo(0, true);
        } else if (e.key === "End") {
          e.preventDefault();
          goTo(n - 1, true);
        }
      });
  
      if (window.matchMedia) {
        var mq = window.matchMedia("(prefers-reduced-motion: reduce)");
        function onMq() {
          reduced = mq.matches;
          if (reduced) {
            clearTimer();
            paused = true;
          } else {
            paused = false;
            scheduleTimer();
          }
          syncToggle();
        }
        if (mq.addEventListener) mq.addEventListener("change", onMq);
        else if (mq.addListener) mq.addListener(onMq);
      }
    }
  
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
    else init();
  })();