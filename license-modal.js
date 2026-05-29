/** 证照预览弹层（门店 / 平台页共用，支持旋转与双指缩放） */
window.LicenseModal = (function () {
  const ROTATE_LEFT_SVG =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M7.11 8.53 5.7 7.11C4.8 8.27 4.24 9.61 4.07 11h2.02c.14-.87.49-1.72 1.02-2.47zM6.09 13H4.07c.17 1.39.72 2.73 1.62 3.89l1.41-1.42c-.52-.75-.87-1.59-1.01-2.47zm1.01 5.32 1.41 1.42c1.04-1.17 1.65-2.61 1.73-4.14H8.14c-.07 1.16-.45 2.25-1.04 3.22zM12 4.07V2.05c-2.01.2-3.86.88-5.32 1.93l1.42 1.42A7.94 7.94 0 0 1 12 4.07zm7.92 2.97-1.41-1.41A7.948 7.948 0 0 0 12 5.09V7.1c1.57.18 3.02.76 4.24 1.58l1.68-1.63zm-2.02 6.98c-.14.87-.49 1.72-1.02 2.47l1.41 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.01zm-1.04 5.32c-.59.97-.97 2.06-1.04 3.22h2.02c.08-1.53.69-2.97 1.73-4.14l-1.41-1.42c-.53.75-.88 1.6-1.02 2.47zM12 19.93v2.02c2.01-.2 3.86-.88 5.32-1.93l-1.42-1.42a7.948 7.948 0 0 1-3.9 1.33zm-5.32-1.93 1.42 1.42A7.948 7.948 0 0 0 12 21.95v-2.02a7.94 7.94 0 0 1-3.9-1.33l-1.42 1.33z"/></svg>';

  const ROTATE_RIGHT_SVG =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M15.55 5.55 14.12 4.12a8.003 8.003 0 0 0-8.07 11.88h2.02c.14-1.16.49-2.01 1.02-2.76l1.41 1.41c-.9 1.16-1.45 2.5-1.62 3.89h2.02c.17-1.39.72-2.73 1.62-3.89l-1.42-1.42c.53-.75.88-1.6 1.02-2.47zm-1.04 5.32c-.14.87-.49 1.72-1.02 2.47l1.41 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.01c-.07 1.16-.45 2.25-1.04 3.22zM12 4.07V2.05c-2.01.2-3.86.88-5.32 1.93l1.42 1.42A7.94 7.94 0 0 1 12 4.07zm7.92 2.97-1.41-1.41A7.948 7.948 0 0 0 12 5.09V7.1c1.57.18 3.02.76 4.24 1.58l1.68-1.63zm-2.02 6.98c-.14.87-.49 1.72-1.02 2.47l1.41 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.01zm-1.04 5.32c-.59.97-.97 2.06-1.04 3.22h2.02c.08-1.53.69-2.97 1.73-4.14l-1.41-1.42c-.53.75-.88 1.6-1.02 2.47zM12 19.93v2.02c2.01-.2 3.86-.88 5.32-1.93l-1.42-1.42a7.948 7.948 0 0 1-3.9 1.33zm-5.32-1.93 1.42 1.42A7.948 7.948 0 0 0 12 21.95v-2.02a7.94 7.94 0 0 1-3.9-1.33l-1.42 1.33z"/></svg>';

  const MIN_SCALE = 1;
  const MAX_SCALE = 4;

  let scrollLockY = 0;
  let imageRotation = 0;
  const blockRotations = {};
  const zoomStates = new WeakMap();

  let modalEl;
  let modalTitleEl;
  let modalTagEl;
  let modalStoreNameEl;
  let licenseImageEl;
  let licenseSingleEl;
  let licenseAllEl;
  let licenseToolbarGlobalEl;
  let licenseRotatorEl;
  let licenseViewportEl;
  let rotateDegreeEl;
  let rotateLeftEl;
  let rotateRightEl;
  let resetRotateEl;

  function lockScroll() {
    scrollLockY = window.scrollY;
    document.body.style.position = "fixed";
    document.body.style.top = `-${scrollLockY}px`;
    document.body.style.left = "0";
    document.body.style.right = "0";
    document.body.style.width = "100%";
  }

  function unlockScroll() {
    document.body.style.position = "";
    document.body.style.top = "";
    document.body.style.left = "";
    document.body.style.right = "";
    document.body.style.width = "";
    window.scrollTo(0, scrollLockY);
  }

  function ensureZoomLayer(rotator) {
    if (!rotator) return null;
    let layer = rotator.querySelector(".license-zoom-layer");
    if (layer) return layer;
    layer = document.createElement("div");
    layer.className = "license-zoom-layer";
    const img = rotator.querySelector(".license-image");
    if (img) layer.appendChild(img);
    rotator.appendChild(layer);
    return layer;
  }

  function getZoomState(viewport) {
    if (!zoomStates.has(viewport)) {
      zoomStates.set(viewport, { scale: 1, x: 0, y: 0 });
    }
    return zoomStates.get(viewport);
  }

  function applyZoom(viewport) {
    const rotator = viewport.querySelector(".license-rotator");
    const layer = ensureZoomLayer(rotator);
    if (!layer) return;
    const { scale, x, y } = getZoomState(viewport);
    layer.style.transform = `translate(${x}px, ${y}px) scale(${scale})`;
    viewport.classList.toggle("is-zoomed", scale > 1.02);
  }

  function resetZoom(viewport) {
    if (!viewport) return;
    const state = getZoomState(viewport);
    state.scale = 1;
    state.x = 0;
    state.y = 0;
    applyZoom(viewport);
  }

  function touchDistance(touches) {
    const dx = touches[0].clientX - touches[1].clientX;
    const dy = touches[0].clientY - touches[1].clientY;
    return Math.hypot(dx, dy);
  }

  function bindPinchZoom(viewport) {
    if (!viewport || viewport.dataset.pinchBound === "1") return;
    viewport.dataset.pinchBound = "1";

    const rotator = viewport.querySelector(".license-rotator");
    ensureZoomLayer(rotator);

    let pinchStartDist = 0;
    let pinchStartScale = 1;
    let panStartX = 0;
    let panStartY = 0;
    let panOriginX = 0;
    let panOriginY = 0;

    viewport.addEventListener(
      "touchstart",
      (e) => {
        if (e.touches.length === 2) {
          pinchStartDist = touchDistance(e.touches);
          pinchStartScale = getZoomState(viewport).scale;
        } else if (e.touches.length === 1 && getZoomState(viewport).scale > 1.02) {
          panStartX = e.touches[0].clientX;
          panStartY = e.touches[0].clientY;
          const state = getZoomState(viewport);
          panOriginX = state.x;
          panOriginY = state.y;
        }
      },
      { passive: true }
    );

    viewport.addEventListener(
      "touchmove",
      (e) => {
        if (e.touches.length === 2) {
          e.preventDefault();
          const state = getZoomState(viewport);
          const dist = touchDistance(e.touches);
          if (pinchStartDist > 0) {
            state.scale = Math.min(
              MAX_SCALE,
              Math.max(MIN_SCALE, pinchStartScale * (dist / pinchStartDist))
            );
            if (state.scale <= 1.02) {
              state.scale = 1;
              state.x = 0;
              state.y = 0;
            }
            applyZoom(viewport);
          }
        } else if (e.touches.length === 1 && getZoomState(viewport).scale > 1.02) {
          e.preventDefault();
          const state = getZoomState(viewport);
          state.x = panOriginX + (e.touches[0].clientX - panStartX);
          state.y = panOriginY + (e.touches[0].clientY - panStartY);
          applyZoom(viewport);
        }
      },
      { passive: false }
    );

    viewport.addEventListener("touchend", () => {
      pinchStartDist = 0;
    });
  }

  function bindAllPinchZoom() {
    if (licenseViewportEl) bindPinchZoom(licenseViewportEl);
    licenseAllEl.querySelectorAll(".license-viewport").forEach(bindPinchZoom);
  }

  function buildBlockToolbarHtml() {
    return `
      <div class="license-toolbar license-toolbar--block">
        <button type="button" class="toolbar-btn" data-action="left" aria-label="逆时针旋转">
          ${ROTATE_LEFT_SVG}左转
        </button>
        <span class="rotate-degree" data-degree>0°</span>
        <button type="button" class="toolbar-btn" data-action="right" aria-label="顺时针旋转">
          ${ROTATE_RIGHT_SVG}右转
        </button>
        <button type="button" class="toolbar-btn toolbar-btn--ghost" data-action="reset" aria-label="重置">重置</button>
      </div>
    `;
  }

  function applySingleRotation() {
    if (!licenseRotatorEl) return;
    licenseRotatorEl.style.transform = `rotate(${imageRotation}deg)`;
    licenseRotatorEl.classList.toggle("is-landscape", imageRotation % 180 !== 0);
    if (rotateDegreeEl) rotateDegreeEl.textContent = `${imageRotation}°`;
  }

  function applyBlockRotation(block) {
    const type = block.dataset.licenseType;
    const deg = blockRotations[type] || 0;
    const rotator = block.querySelector(".license-rotator");
    const degreeEl = block.querySelector("[data-degree]");
    if (rotator) {
      rotator.style.transform = `rotate(${deg}deg)`;
      rotator.classList.toggle("is-landscape", deg % 180 !== 0);
    }
    if (degreeEl) degreeEl.textContent = `${deg}°`;
  }

  function bindBlockToolbar(block) {
    const type = block.dataset.licenseType;
    const viewport = block.querySelector(".license-viewport");
    bindPinchZoom(viewport);

    block.querySelectorAll("[data-action]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const action = btn.dataset.action;
        if (action === "left") {
          blockRotations[type] = ((blockRotations[type] || 0) - 90 + 360) % 360;
        } else if (action === "right") {
          blockRotations[type] = ((blockRotations[type] || 0) + 90) % 360;
        } else if (action === "reset") {
          blockRotations[type] = 0;
          resetZoom(viewport);
        }
        applyBlockRotation(block);
      });
    });

    const img = block.querySelector(".license-image");
    if (img) {
      img.onload = () => {
        applyBlockRotation(block);
        resetZoom(viewport);
      };
    }
  }

  function setLicenseMode(mode) {
    const isAll = mode === "all";
    licenseSingleEl.classList.toggle("hidden", isAll);
    licenseAllEl.classList.toggle("hidden", !isAll);
    if (licenseToolbarGlobalEl) {
      licenseToolbarGlobalEl.classList.toggle("hidden", isAll);
    }
  }

  function resetSingleView() {
    resetRotation();
    resetZoom(licenseViewportEl);
  }

  function resetRotation() {
    imageRotation = 0;
    applySingleRotation();
  }

  function resetBlockRotations(types) {
    for (const key of Object.keys(blockRotations)) {
      delete blockRotations[key];
    }
    for (const type of types) {
      blockRotations[type] = 0;
    }
  }

  function buildAllLicensesHtml(blocks) {
    return blocks
      .map(
        (block) => `
        <section class="license-block" data-license-type="${block.type}">
          <h4 class="license-block__title">${block.label}</h4>
          <div class="license-viewport license-viewport--stack">
            <div class="license-rotator">
              <div class="license-zoom-layer">
                <img class="license-image" src="${block.src}" alt="${block.alt}" />
              </div>
            </div>
          </div>
          ${buildBlockToolbarHtml()}
        </section>
      `
      )
      .join("");
  }

  function openSingle(options) {
    const { tag, title, subtitle, src, alt } = options;
    if (modalTagEl) modalTagEl.textContent = tag || "证照公示";
    modalTitleEl.textContent = title;
    modalStoreNameEl.textContent = subtitle;
    resetSingleView();
    setLicenseMode("single");
    licenseAllEl.innerHTML = "";
    ensureZoomLayer(licenseRotatorEl);
    licenseImageEl.onload = () => {
      applySingleRotation();
      resetZoom(licenseViewportEl);
    };
    licenseImageEl.src = src;
    licenseImageEl.alt = alt || title;
    modalEl.classList.remove("hidden");
    lockScroll();
    bindPinchZoom(licenseViewportEl);
  }

  function openMultiple(options) {
    const { tag, title, subtitle, blocks } = options;
    if (modalTagEl) modalTagEl.textContent = tag || "全部证照";
    modalTitleEl.textContent = title;
    modalStoreNameEl.textContent = subtitle;
    resetBlockRotations(blocks.map((b) => b.type));
    resetRotation();
    setLicenseMode("all");
    licenseAllEl.innerHTML = buildAllLicensesHtml(blocks);
    licenseAllEl.querySelectorAll(".license-block").forEach((block) => {
      bindBlockToolbar(block);
      applyBlockRotation(block);
    });
    licenseImageEl.onload = null;
    licenseImageEl.src = "";
    modalEl.classList.remove("hidden");
    lockScroll();
    bindAllPinchZoom();
  }

  function close() {
    modalEl.classList.add("hidden");
    licenseImageEl.onload = null;
    licenseImageEl.src = "";
    licenseAllEl.innerHTML = "";
    setLicenseMode("single");
    resetSingleView();
    unlockScroll();
  }

  function init() {
    modalEl = document.getElementById("licenseModal");
    modalTitleEl = document.getElementById("modalTitle");
    modalTagEl = document.getElementById("modalTag");
    modalStoreNameEl = document.getElementById("modalStoreName");
    licenseImageEl = document.getElementById("licenseImage");
    licenseSingleEl = document.getElementById("licenseSingle");
    licenseAllEl = document.getElementById("licenseAll");
    licenseToolbarGlobalEl = document.getElementById("licenseToolbarGlobal");
    licenseRotatorEl = document.getElementById("licenseRotator");
    licenseViewportEl = document.getElementById("licenseViewport");
    rotateDegreeEl = document.getElementById("rotateDegree");
    rotateLeftEl = document.getElementById("rotateLeft");
    rotateRightEl = document.getElementById("rotateRight");
    resetRotateEl = document.getElementById("resetRotate");

    ensureZoomLayer(licenseRotatorEl);
    bindPinchZoom(licenseViewportEl);

    if (rotateLeftEl) {
      rotateLeftEl.addEventListener("click", () => {
        imageRotation = ((imageRotation - 90) % 360 + 360) % 360;
        applySingleRotation();
      });
    }
    if (rotateRightEl) {
      rotateRightEl.addEventListener("click", () => {
        imageRotation = ((imageRotation + 90) % 360 + 360) % 360;
        applySingleRotation();
      });
    }
    if (resetRotateEl) {
      resetRotateEl.addEventListener("click", resetSingleView);
    }

    modalEl.querySelectorAll("[data-close]").forEach((el) => {
      el.addEventListener("click", close);
    });

    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && !modalEl.classList.contains("hidden")) {
        close();
      }
    });
  }

  return { init, openSingle, openMultiple, close };
})();
