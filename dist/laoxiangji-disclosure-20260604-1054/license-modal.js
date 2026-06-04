/** 证照预览弹层（门店 / 平台页共用，支持旋转与双指缩放） */
window.LicenseModal = (function () {
  const MIN_SCALE = 1;
  const MAX_SCALE = 4;

  let scrollLockY = 0;
  let imageRotation = 0;
  let imageRotationApplied = 0;
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
    layer.style.transform = `translate3d(${x}px, ${y}px, 0) scale(${scale})`;
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

  function setViewportLoading(viewport, loading) {
    if (!viewport) return;
    viewport.classList.toggle("is-loading", loading);
  }

  function syncImageOrientationClass(imgEl) {
    if (!imgEl) return;
    imgEl.classList.remove("license-image--portrait");
    if (imgEl.naturalHeight > imgEl.naturalWidth) {
      imgEl.classList.add("license-image--portrait");
    }
  }

  function touchDistance(touches) {
    const dx = touches[0].clientX - touches[1].clientX;
    const dy = touches[0].clientY - touches[1].clientY;
    return Math.hypot(dx, dy);
  }

  function touchCenter(touches) {
    return {
      x: (touches[0].clientX + touches[1].clientX) / 2,
      y: (touches[0].clientY + touches[1].clientY) / 2,
    };
  }

  function setTouching(viewport, active) {
    viewport.classList.toggle("is-touching", active);
  }

  function bindPinchZoom(viewport) {
    if (!viewport || viewport.dataset.pinchBound === "1") return;
    viewport.dataset.pinchBound = "1";

    const rotator = viewport.querySelector(".license-rotator");
    ensureZoomLayer(rotator);

    let gesture = null;
    let pinchStartDist = 0;
    let pinchStartScale = 1;
    let pinchStartCenterX = 0;
    let pinchStartCenterY = 0;
    let pinchStartPanX = 0;
    let pinchStartPanY = 0;
    let panPointerX = 0;
    let panPointerY = 0;
    let panOriginX = 0;
    let panOriginY = 0;

    function beginPinch(touches) {
      gesture = "pinch";
      setTouching(viewport, true);
      const state = getZoomState(viewport);
      pinchStartDist = touchDistance(touches);
      pinchStartScale = state.scale;
      const center = touchCenter(touches);
      pinchStartCenterX = center.x;
      pinchStartCenterY = center.y;
      pinchStartPanX = state.x;
      pinchStartPanY = state.y;
    }

    function beginPan(touch) {
      gesture = "pan";
      setTouching(viewport, true);
      panPointerX = touch.clientX;
      panPointerY = touch.clientY;
      const state = getZoomState(viewport);
      panOriginX = state.x;
      panOriginY = state.y;
    }

    function endGestureIfNeeded(touches) {
      if (touches.length > 0) return;
      const state = getZoomState(viewport);
      if (state.scale <= 1.02) {
        state.scale = 1;
        state.x = 0;
        state.y = 0;
        applyZoom(viewport);
      }
      gesture = null;
      pinchStartDist = 0;
      setTouching(viewport, false);
    }

    viewport.addEventListener(
      "touchstart",
      (e) => {
        if (e.touches.length >= 2) {
          beginPinch(e.touches);
        } else if (
          e.touches.length === 1 &&
          !gesture &&
          getZoomState(viewport).scale > 1.02
        ) {
          beginPan(e.touches[0]);
        }
      },
      { passive: true }
    );

    viewport.addEventListener(
      "touchmove",
      (e) => {
        if (gesture === "pinch" && e.touches.length >= 2) {
          e.preventDefault();
          const state = getZoomState(viewport);
          const dist = touchDistance(e.touches);
          if (pinchStartDist > 0) {
            state.scale = Math.min(
              MAX_SCALE,
              Math.max(MIN_SCALE, pinchStartScale * (dist / pinchStartDist))
            );
          }
          const center = touchCenter(e.touches);
          state.x = pinchStartPanX + (center.x - pinchStartCenterX);
          state.y = pinchStartPanY + (center.y - pinchStartCenterY);
          applyZoom(viewport);
          return;
        }

        if (
          gesture === "pan" &&
          e.touches.length === 1 &&
          getZoomState(viewport).scale > 1.02
        ) {
          e.preventDefault();
          const state = getZoomState(viewport);
          state.x = panOriginX + (e.touches[0].clientX - panPointerX);
          state.y = panOriginY + (e.touches[0].clientY - panPointerY);
          applyZoom(viewport);
        }
      },
      { passive: false }
    );

    viewport.addEventListener("touchend", (e) => {
      endGestureIfNeeded(e.touches);
    });

    viewport.addEventListener("touchcancel", (e) => {
      endGestureIfNeeded(e.touches);
    });
  }

  function bindAllPinchZoom() {
    if (licenseViewportEl) bindPinchZoom(licenseViewportEl);
    licenseAllEl.querySelectorAll(".license-viewport").forEach(bindPinchZoom);
  }

  function normalizeRotation(deg) {
    return ((deg % 360) + 360) % 360;
  }

  function buildBlockToolbarHtml() {
    return `
      <div class="license-toolbar license-toolbar--block">
        <button type="button" class="toolbar-btn" data-action="left" aria-label="逆时针旋转">左转</button>
        <span class="rotate-degree" data-degree>0°</span>
        <button type="button" class="toolbar-btn" data-action="right" aria-label="顺时针旋转">右转</button>
        <button type="button" class="toolbar-btn toolbar-btn--ghost" data-action="reset" aria-label="重置">重置</button>
      </div>
    `;
  }

  function applySingleRotation() {
    if (!licenseRotatorEl) return;
    licenseRotatorEl.style.transform = `rotate(${imageRotationApplied}deg)`;
    licenseRotatorEl.classList.toggle("is-landscape", imageRotation % 180 !== 0);
    if (rotateDegreeEl) rotateDegreeEl.textContent = `${imageRotation}°`;
  }

  function applyBlockRotation(block) {
    const type = block.dataset.licenseType;
    const applied = blockRotations[type] || 0;
    const deg = normalizeRotation(applied);
    const rotator = block.querySelector(".license-rotator");
    const degreeEl = block.querySelector("[data-degree]");
    if (rotator) {
      rotator.style.transform = `rotate(${applied}deg)`;
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
          blockRotations[type] = (blockRotations[type] || 0) - 90;
        } else if (action === "right") {
          blockRotations[type] = (blockRotations[type] || 0) + 90;
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
        syncImageOrientationClass(img);
        setViewportLoading(viewport, false);
        applyBlockRotation(block);
        resetZoom(viewport);
      };
      img.onerror = () => {
        img.classList.remove("license-image--portrait");
        setViewportLoading(viewport, false);
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
    imageRotationApplied = 0;
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
          <div class="license-viewport license-viewport--stack is-loading">
            <div class="license-rotator">
              <div class="license-zoom-layer">
                <img class="license-image" data-src="${block.src}" alt="${block.alt}" />
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
    setViewportLoading(licenseViewportEl, true);
    licenseImageEl.classList.remove("license-image--portrait");
    ensureZoomLayer(licenseRotatorEl);
    licenseImageEl.onload = () => {
      syncImageOrientationClass(licenseImageEl);
      setViewportLoading(licenseViewportEl, false);
      applySingleRotation();
      resetZoom(licenseViewportEl);
    };
    licenseImageEl.onerror = () => {
      licenseImageEl.classList.remove("license-image--portrait");
      setViewportLoading(licenseViewportEl, false);
    };
    licenseImageEl.src = "";
    licenseImageEl.alt = alt || title;
    modalEl.classList.remove("hidden");
    lockScroll();
    bindPinchZoom(licenseViewportEl);
    requestAnimationFrame(() => {
      licenseImageEl.src = src;
    });
  }

  function openMultiple(options) {
    const { tag, title, subtitle, blocks } = options;
    if (modalTagEl) modalTagEl.textContent = tag || "全部证照";
    modalTitleEl.textContent = title;
    modalStoreNameEl.textContent = subtitle;
    resetBlockRotations(blocks.map((b) => b.type));
    resetRotation();
    setLicenseMode("all");
    modalEl.classList.remove("hidden");
    lockScroll();
    licenseAllEl.innerHTML = buildAllLicensesHtml(blocks);
    licenseAllEl.querySelectorAll(".license-block").forEach((block) => {
      const viewport = block.querySelector(".license-viewport");
      setViewportLoading(viewport, true);
      bindBlockToolbar(block);
      applyBlockRotation(block);
      const img = block.querySelector(".license-image");
      if (img) {
        img.src = img.dataset.src || "";
      }
    });
    licenseImageEl.onload = null;
    licenseImageEl.onerror = null;
    licenseImageEl.src = "";
    bindAllPinchZoom();
  }

  function close() {
    modalEl.classList.add("hidden");
    licenseImageEl.onload = null;
    licenseImageEl.onerror = null;
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
        imageRotationApplied -= 90;
        imageRotation = normalizeRotation(imageRotationApplied);
        applySingleRotation();
      });
    }
    if (rotateRightEl) {
      rotateRightEl.addEventListener("click", () => {
        imageRotationApplied += 90;
        imageRotation = normalizeRotation(imageRotationApplied);
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
