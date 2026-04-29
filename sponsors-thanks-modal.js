/**
 * Modal dos patrocinadores: mensagem geral + vitrine com todos os logos.
 */
(function () {
  "use strict";

  function $(id) {
    return document.getElementById(id);
  }

  function buildLogosWall(wall, list) {
    wall.innerHTML = "";
    var tiles = list.querySelectorAll(".sponsor-tile");

    for (var i = 0; i < tiles.length; i++) {
      var tile = tiles[i];
      var img = tile.querySelector(".sponsor-tile__img");
      var nameEl = tile.querySelector(".sponsor-tile__name");
      if (!img || !nameEl) continue;

      var item = document.createElement("div");
      item.className = "sponsors-modal__logos-item";

      var im = document.createElement("img");
      im.className = "sponsors-modal__logos-img";
      im.src = img.getAttribute("src") || "";
      im.alt = img.getAttribute("alt") || nameEl.textContent.trim();
      im.decoding = "async";
      im.loading = "lazy";

      var cls = img.className || "";
      if (
        cls.indexOf("sponsor-tile__img--mundo-pet") !== -1 ||
        cls.indexOf("sponsor-tile__img--gean-veiculos") !== -1 ||
        cls.indexOf("sponsor-tile__img--marcelo-personal") !== -1
      ) {
        im.classList.add("sponsors-modal__logos-img--tall");
      }
      if (cls.indexOf("sponsor-tile__img--marcelo-personal") !== -1) {
        item.classList.add("sponsors-modal__logos-item--dark");
      }

      var caption = document.createElement("p");
      caption.className = "sponsors-modal__logos-name";
      caption.textContent = nameEl.textContent.trim();

      item.appendChild(im);
      item.appendChild(caption);
      wall.appendChild(item);
    }
  }

  function init() {
    var modal = $("sponsors-thanks-modal");
    var openBtn = $("sponsors-open-thanks-modal");
    var wall = $("sponsors-modal-logos-wall");
    var list = $("sponsors-showcase-list");
    if (!modal || !openBtn || !wall || !list) return;

    var tiles = list.querySelectorAll(".sponsor-tile");
    if (tiles.length === 0) {
      openBtn.hidden = true;
      return;
    }

    buildLogosWall(wall, list);
    openBtn.hidden = false;

    var panel = modal.querySelector(".sponsors-modal__panel");
    var backdrop = modal.querySelector(".sponsors-modal__backdrop");
    var closeBtn = modal.querySelector(".sponsors-modal__close");

    var prevActive = null;

    function openModal() {
      prevActive = document.activeElement;
      modal.hidden = false;
      document.body.classList.add("sponsors-modal-is-open");
      if (closeBtn) closeBtn.focus();
    }

    function closeModal() {
      modal.hidden = true;
      document.body.classList.remove("sponsors-modal-is-open");
      if (prevActive && typeof prevActive.focus === "function") {
        prevActive.focus();
      }
    }

    openBtn.addEventListener("click", openModal);
    if (closeBtn) closeBtn.addEventListener("click", closeModal);
    if (backdrop) backdrop.addEventListener("click", closeModal);

    document.addEventListener("keydown", function (e) {
      if (!modal.hidden && e.key === "Escape") {
        e.preventDefault();
        closeModal();
      }
    });

    if (panel) {
      panel.addEventListener("keydown", function (e) {
        if (e.key !== "Tab" || modal.hidden) return;
        var focusables = panel.querySelectorAll(
          'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
        );
        var nodes = [];
        for (var j = 0; j < focusables.length; j++) {
          if (!focusables[j].hasAttribute("disabled") && !focusables[j].hidden) {
            nodes.push(focusables[j]);
          }
        }
        if (nodes.length === 0) return;
        var first = nodes[0];
        var last = nodes[nodes.length - 1];
        if (e.shiftKey) {
          if (document.activeElement === first) {
            e.preventDefault();
            last.focus();
          }
        } else {
          if (document.activeElement === last) {
            e.preventDefault();
            first.focus();
          }
        }
      });
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
