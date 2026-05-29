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
  const licenseRotatorEl = document.getElementById("licenseRotator");
  const licenseViewportEl = document.getElementById("licenseViewport");
  const rotateDegreeEl = document.getElementById("rotateDegree");
  const rotateLeftEl = document.getElementById("rotateLeft");
  const rotateRightEl = document.getElementById("rotateRight");
  const resetRotateEl = document.getElementById("resetRotate");

  let imageRotation = 0;

  const LICENSE_LABELS = {
    business: "营业执照",
    food: "食品经营许可证",
  };

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
      <div class="store-card__head">
        <span class="city-tag">${escapeHtml(store.city)}</span>
      </div>
      <h3 class="store-name">${escapeHtml(store.name)}</h3>
      <div class="store-licenses">
        <button type="button" class="license-chip" data-license="business">
          ${DOC_ICON}营业执照
        </button>
        <button type="button" class="license-chip" data-license="food">
          ${DOC_ICON}食品经营许可证
        </button>
      </div>
    `;

    li.querySelectorAll(".license-chip").forEach((btn) => {
      btn.addEventListener("click", () => {
        const type = btn.dataset.license;
        openLicenseModal(store, type);
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

  function applyRotation() {
    if (!licenseRotatorEl) return;
    licenseRotatorEl.style.transform = `rotate(${imageRotation}deg)`;
    const isSideways = imageRotation % 180 !== 0;
    licenseRotatorEl.classList.toggle("is-landscape", isSideways);
    if (rotateDegreeEl) rotateDegreeEl.textContent = `${imageRotation}°`;
    if (licenseViewportEl) licenseViewportEl.scrollTop = 0;
  }

  function setRotation(deg) {
    imageRotation = ((deg % 360) + 360) % 360;
    applyRotation();
  }

  function resetRotation() {
    setRotation(0);
  }

  function openLicenseModal(store, type) {
    const label = LICENSE_LABELS[type];
    const src = store.licenses[type];
    if (modalTagEl) modalTagEl.textContent = "证照公示";
    modalTitleEl.textContent = label;
    modalStoreNameEl.textContent = store.name;
    resetRotation();
    licenseImageEl.onload = applyRotation;
    licenseImageEl.src = src;
    licenseImageEl.alt = `${store.name} ${label}`;
    modalEl.classList.remove("hidden");
    lockScroll();
  }

  function closeModal() {
    modalEl.classList.add("hidden");
    licenseImageEl.onload = null;
    licenseImageEl.src = "";
    resetRotation();
    unlockScroll();
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
