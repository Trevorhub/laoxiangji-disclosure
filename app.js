(function () {
  const provinceEl = document.getElementById("province");
  const cityEl = document.getElementById("city");
  const keywordEl = document.getElementById("keyword");
  const clearKeywordEl = document.getElementById("clearKeyword");
  const storeListEl = document.getElementById("storeList");
  const emptyStateEl = document.getElementById("emptyState");
  const resultCountEl = document.getElementById("resultCount");
  const resetFiltersEl = document.getElementById("resetFilters");
  const modalEl = document.getElementById("licenseModal");
  const modalTitleEl = document.getElementById("modalTitle");
  const modalTagEl = document.getElementById("modalTag");
  const modalStoreNameEl = document.getElementById("modalStoreName");
  const licenseImageEl = document.getElementById("licenseImage");
  const licenseSingleEl = document.getElementById("licenseSingle");
  const licenseAllEl = document.getElementById("licenseAll");
  const licenseToolbarGlobalEl = document.getElementById("licenseToolbarGlobal");
  const licenseRotatorEl = document.getElementById("licenseRotator");
  const licenseViewportEl = document.getElementById("licenseViewport");
  const rotateDegreeEl = document.getElementById("rotateDegree");
  const rotateLeftEl = document.getElementById("rotateLeft");
  const rotateRightEl = document.getElementById("rotateRight");
  const resetRotateEl = document.getElementById("resetRotate");

  let imageRotation = 0;
  const blockRotations = { business: 0, food: 0 };

  const LICENSE_LABELS = {
    business: "营业执照",
    food: "食品经营许可证",
  };

  const ROTATE_LEFT_SVG =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M7.11 8.53 5.7 7.11C4.8 8.27 4.24 9.61 4.07 11h2.02c.14-.87.49-1.72 1.02-2.47zM6.09 13H4.07c.17 1.39.72 2.73 1.62 3.89l1.41-1.42c-.52-.75-.87-1.59-1.01-2.47zm1.01 5.32 1.41 1.42c1.04-1.17 1.65-2.61 1.73-4.14H8.14c-.07 1.16-.45 2.25-1.04 3.22zM12 4.07V2.05c-2.01.2-3.86.88-5.32 1.93l1.42 1.42A7.94 7.94 0 0 1 12 4.07zm7.92 2.97-1.41-1.41A7.948 7.948 0 0 0 12 5.09V7.1c1.57.18 3.02.76 4.24 1.58l1.68-1.63zm-2.02 6.98c-.14.87-.49 1.72-1.02 2.47l1.41 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.01zm-1.04 5.32c-.59.97-.97 2.06-1.04 3.22h2.02c.08-1.53.69-2.97 1.73-4.14l-1.41-1.42c-.53.75-.88 1.6-1.02 2.47zM12 19.93v2.02c2.01-.2 3.86-.88 5.32-1.93l-1.42-1.42a7.948 7.948 0 0 1-3.9 1.33zm-5.32-1.93 1.42 1.42A7.948 7.948 0 0 0 12 21.95v-2.02a7.94 7.94 0 0 1-3.9-1.33l-1.42 1.33z"/></svg>';

  const ROTATE_RIGHT_SVG =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M15.55 5.55 14.12 4.12a8.003 8.003 0 0 0-8.07 11.88h2.02c.14-1.16.49-2.01 1.02-2.76l1.41 1.41c-.9 1.16-1.45 2.5-1.62 3.89h2.02c.17-1.39.72-2.73 1.62-3.89l-1.42-1.42c.53-.75.88-1.6 1.02-2.47zm-1.04 5.32c-.14.87-.49 1.72-1.02 2.47l1.41 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.01c-.07 1.16-.45 2.25-1.04 3.22zM12 4.07V2.05c-2.01.2-3.86.88-5.32 1.93l1.42 1.42A7.94 7.94 0 0 1 12 4.07zm7.92 2.97-1.41-1.41A7.948 7.948 0 0 0 12 5.09V7.1c1.57.18 3.02.76 4.24 1.58l1.68-1.63zm-2.02 6.98c-.14.87-.49 1.72-1.02 2.47l1.41 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.01zm-1.04 5.32c-.59.97-.97 2.06-1.04 3.22h2.02c.08-1.53.69-2.97 1.73-4.14l-1.41-1.42c-.53.75-.88 1.6-1.02 2.47zM12 19.93v2.02c2.01-.2 3.86-.88 5.32-1.93l-1.42-1.42a7.948 7.948 0 0 1-3.9 1.33zm-5.32-1.93 1.42 1.42A7.948 7.948 0 0 0 12 21.95v-2.02a7.94 7.94 0 0 1-3.9-1.33l-1.42 1.33z"/></svg>';

  function buildRegionMaps() {
    const provinces = new Map();
    for (const store of STORES) {
      if (!provinces.has(store.province)) {
        provinces.set(store.province, new Set());
      }
      provinces.get(store.province).add(store.city);
    }
    return provinces;
  }

  const regionMap = buildRegionMaps();

  function fillProvinceOptions() {
    const sorted = [...regionMap.keys()].sort((a, b) => a.localeCompare(b, "zh"));
    for (const p of sorted) {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      provinceEl.appendChild(opt);
    }
  }

  function fillCityOptions(province) {
    cityEl.innerHTML = '<option value="">全部</option>';
    if (!province) {
      cityEl.disabled = true;
      return;
    }
    const cities = [...regionMap.get(province)].sort((a, b) => a.localeCompare(b, "zh"));
    for (const c of cities) {
      const opt = document.createElement("option");
      opt.value = c;
      opt.textContent = c;
      cityEl.appendChild(opt);
    }
    cityEl.disabled = false;
  }

  function getFilteredStores() {
    const province = provinceEl.value;
    const city = cityEl.value;
    const keyword = keywordEl.value.trim().toLowerCase();

    return STORES.filter((store) => {
      if (province && store.province !== province) return false;
      if (city && store.city !== city) return false;
      if (keyword && !store.name.toLowerCase().includes(keyword)) return false;
      return true;
    });
  }

  const DOC_ICON =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm-1 2.5L18.5 9H13V4.5zM8 13h8v1.5H8V13zm0 3.5h5V18H8v-1.5z"/></svg>';

  function renderStoreItem(store) {
    const li = document.createElement("li");
    li.className = "store-card";
    li.innerHTML = `
      <div class="store-card__body" role="button" tabindex="0" aria-label="查看${escapeHtml(store.name)}全部证照">
        <div class="store-card__head">
          <span class="city-tag">${escapeHtml(store.city)}</span>
        </div>
        <h3 class="store-name">${escapeHtml(store.name)}</h3>
      </div>
      <div class="store-licenses">
        <button type="button" class="license-chip" data-license="business">
          ${DOC_ICON}营业执照
        </button>
        <button type="button" class="license-chip" data-license="food">
          ${DOC_ICON}食品经营许可证
        </button>
      </div>
    `;

    const cardBody = li.querySelector(".store-card__body");
    cardBody.addEventListener("click", () => openAllLicensesModal(store));
    cardBody.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openAllLicensesModal(store);
      }
    });

    li.querySelectorAll(".license-chip").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        openLicenseModal(store, btn.dataset.license);
      });
    });

    return li;
  }

  function renderList() {
    const list = getFilteredStores();
    storeListEl.innerHTML = "";

    if (list.length === 0) {
      emptyStateEl.classList.remove("hidden");
      storeListEl.classList.add("hidden");
      resultCountEl.textContent = "";
      return;
    }

    emptyStateEl.classList.add("hidden");
    storeListEl.classList.remove("hidden");
    resultCountEl.innerHTML = `共找到 <strong>${list.length}</strong> 家门店`;

    const frag = document.createDocumentFragment();
    for (const store of list) {
      frag.appendChild(renderStoreItem(store));
    }
    storeListEl.appendChild(frag);
  }

  let scrollLockY = 0;

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

  const LICENSE_TYPES = ["business", "food"];

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
        <button type="button" class="toolbar-btn toolbar-btn--ghost" data-action="reset" aria-label="重置旋转">重置</button>
      </div>
    `;
  }

  function applySingleRotation() {
    if (!licenseRotatorEl) return;
    licenseRotatorEl.style.transform = `rotate(${imageRotation}deg)`;
    const isSideways = imageRotation % 180 !== 0;
    licenseRotatorEl.classList.toggle("is-landscape", isSideways);
    if (rotateDegreeEl) rotateDegreeEl.textContent = `${imageRotation}°`;
    if (licenseViewportEl) licenseViewportEl.scrollTop = 0;
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
    block.querySelectorAll("[data-action]").forEach((btn) => {
      btn.addEventListener("click", () => {
        let deg = blockRotations[type] || 0;
        const action = btn.dataset.action;
        if (action === "left") deg -= 90;
        else if (action === "right") deg += 90;
        else if (action === "reset") deg = 0;
        blockRotations[type] = ((deg % 360) + 360) % 360;
        applyBlockRotation(block);
      });
    });
    const img = block.querySelector(".license-image");
    if (img) img.onload = () => applyBlockRotation(block);
  }

  function resetBlockRotations() {
    blockRotations.business = 0;
    blockRotations.food = 0;
  }

  function setLicenseMode(mode) {
    const isAll = mode === "all";
    licenseSingleEl.classList.toggle("hidden", isAll);
    licenseAllEl.classList.toggle("hidden", !isAll);
    if (licenseToolbarGlobalEl) {
      licenseToolbarGlobalEl.classList.toggle("hidden", isAll);
    }
  }

  function buildAllLicensesHtml(store) {
    return LICENSE_TYPES.map((type) => {
      const label = LICENSE_LABELS[type];
      const src = store.licenses[type];
      return `
        <section class="license-block" data-license-type="${type}">
          <h4 class="license-block__title">${label}</h4>
          <div class="license-viewport license-viewport--stack">
            <div class="license-rotator">
              <img class="license-image" src="${src}" alt="${escapeHtml(store.name)} ${label}" />
            </div>
          </div>
          ${buildBlockToolbarHtml()}
        </section>
      `;
    }).join("");
  }

  function openAllLicensesModal(store) {
    if (modalTagEl) modalTagEl.textContent = "全部证照";
    modalTitleEl.textContent = "门店证照公示";
    modalStoreNameEl.textContent = store.name;
    resetBlockRotations();
    resetRotation();
    setLicenseMode("all");
    licenseAllEl.innerHTML = buildAllLicensesHtml(store);
    licenseAllEl.querySelectorAll(".license-block").forEach((block) => {
      bindBlockToolbar(block);
      applyBlockRotation(block);
    });
    licenseImageEl.onload = null;
    licenseImageEl.src = "";
    modalEl.classList.remove("hidden");
    lockScroll();
  }

  function openLicenseModal(store, type) {
    const label = LICENSE_LABELS[type];
    const src = store.licenses[type];
    if (modalTagEl) modalTagEl.textContent = "证照公示";
    modalTitleEl.textContent = label;
    modalStoreNameEl.textContent = store.name;
    resetRotation();
    setLicenseMode("single");
    licenseAllEl.innerHTML = "";
    licenseImageEl.onload = applySingleRotation;
    licenseImageEl.src = src;
    licenseImageEl.alt = `${store.name} ${label}`;
    modalEl.classList.remove("hidden");
    lockScroll();
  }

  function closeModal() {
    modalEl.classList.add("hidden");
    licenseImageEl.onload = null;
    licenseImageEl.src = "";
    licenseAllEl.innerHTML = "";
    setLicenseMode("single");
    resetRotation();
    resetBlockRotations();
    unlockScroll();
  }

  function setRotation(deg) {
    imageRotation = ((deg % 360) + 360) % 360;
    applySingleRotation();
  }

  function resetRotation() {
    setRotation(0);
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function resetFilters() {
    provinceEl.value = "";
    fillCityOptions("");
    cityEl.value = "";
    keywordEl.value = "";
    clearKeywordEl.classList.add("hidden");
    renderList();
  }

  provinceEl.addEventListener("change", () => {
    fillCityOptions(provinceEl.value);
    cityEl.value = "";
    renderList();
  });

  cityEl.addEventListener("change", renderList);

  keywordEl.addEventListener("input", () => {
    clearKeywordEl.classList.toggle("hidden", !keywordEl.value);
    renderList();
  });

  clearKeywordEl.addEventListener("click", () => {
    keywordEl.value = "";
    clearKeywordEl.classList.add("hidden");
    keywordEl.focus();
    renderList();
  });

  resetFiltersEl.addEventListener("click", resetFilters);

  if (rotateLeftEl) {
    rotateLeftEl.addEventListener("click", () => setRotation(imageRotation - 90));
  }
  if (rotateRightEl) {
    rotateRightEl.addEventListener("click", () => setRotation(imageRotation + 90));
  }
  if (resetRotateEl) {
    resetRotateEl.addEventListener("click", resetRotation);
  }

  modalEl.querySelectorAll("[data-close]").forEach((el) => {
    el.addEventListener("click", closeModal);
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !modalEl.classList.contains("hidden")) {
      closeModal();
    }
  });

  fillProvinceOptions();
  fillCityOptions("");
  renderList();
})();
