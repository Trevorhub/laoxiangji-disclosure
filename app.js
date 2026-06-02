(function () {
  const modal = window.LicenseModal;
  modal.init();

  const provinceEl = document.getElementById("province");
  const cityEl = document.getElementById("city");
  const keywordEl = document.getElementById("keyword");
  const clearKeywordEl = document.getElementById("clearKeyword");
  const storeListEl = document.getElementById("storeList");
  const emptyStateEl = document.getElementById("emptyState");
  const resultCountEl = document.getElementById("resultCount");
  const resetFiltersEl = document.getElementById("resetFilters");

  const LICENSE_LABELS = {
    business: "营业执照",
    food: "食品经营许可证",
  };

  const LICENSE_TYPES = ["business", "food"];

  const DOC_ICON =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm-1 2.5L18.5 9H13V4.5zM8 13h8v1.5H8V13zm0 3.5h5V18H8v-1.5z"/></svg>';

  function getAvailableLicenseTypes(store) {
    return LICENSE_TYPES.filter((type) => !!store.licenses[type]);
  }

  const storeById = new Map(
    STORES.map((store) => [String(store.id).toLowerCase(), store])
  );

  let deepLinkTargetId = null;

  function parseStoreIdFromUrl() {
    const params = new URLSearchParams(window.location.search);
    const id = params.get("storeId") || params.get("store_id");
    return id ? String(id).trim() : "";
  }

  function findStoreById(id) {
    if (!id) return null;
    return storeById.get(String(id).toLowerCase()) || null;
  }

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
    }).sort((a, b) => a.name.localeCompare(b.name, "zh-Hans-CN-u-co-pinyin"));
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function buildLicenseBlocks(store) {
    return getAvailableLicenseTypes(store).map((type) => ({
      type,
      label: LICENSE_LABELS[type],
      src: store.licenses[type],
      alt: `${store.name} ${LICENSE_LABELS[type]}`,
    }));
  }

  function openAllLicensesModal(store) {
    const blocks = buildLicenseBlocks(store);
    if (blocks.length === 0) return;
    modal.openMultiple({
      tag: "门店资质",
      title: "门店证照公示",
      subtitle: store.name,
      blocks,
    });
  }

  function openLicenseModal(store, type) {
    if (!store.licenses[type]) return;
    const label = LICENSE_LABELS[type];
    modal.openSingle({
      tag: "门店资质",
      title: label,
      subtitle: store.name,
      src: store.licenses[type],
      alt: `${store.name} ${label}`,
    });
  }

  function scrollToStoreCard(storeId) {
    const card = storeListEl.querySelector(
      `[data-store-id="${CSS.escape(String(storeId))}"]`
    );
    if (!card) return;
    card.classList.add("is-target");
    requestAnimationFrame(() => {
      card.scrollIntoView({ behavior: "smooth", block: "center" });
    });
    window.setTimeout(() => {
      card.classList.remove("is-target");
    }, 3200);
  }

  function renderStoreItem(store) {
    const availableTypes = getAvailableLicenseTypes(store);
    const li = document.createElement("li");
    li.className = "store-card";
    li.dataset.storeId = String(store.id);
    li.innerHTML = `
      <div class="store-card__body" role="button" tabindex="0" aria-label="查看${escapeHtml(store.name)}全部证照">
        <div class="store-card__head">
          <span class="city-tag">${escapeHtml(store.city)}</span>
        </div>
        <h3 class="store-name">${escapeHtml(store.name)}</h3>
      </div>
      <div class="store-licenses">
        ${availableTypes
          .map(
            (type) => `
          <button type="button" class="license-chip" data-license="${type}">
            ${DOC_ICON}${LICENSE_LABELS[type]}
          </button>
        `
          )
          .join("")}
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

  function renderList(options) {
    const scrollToId = options && options.scrollToId;
    const list = getFilteredStores();
    storeListEl.innerHTML = "";

    if (list.length === 0) {
      emptyStateEl.classList.remove("hidden");
      storeListEl.classList.add("hidden");
      resultCountEl.textContent = "";
      return false;
    }

    emptyStateEl.classList.add("hidden");
    storeListEl.classList.remove("hidden");

    const targetStore =
      scrollToId && list.find((store) => String(store.id) === String(scrollToId));
    if (targetStore) {
      resultCountEl.innerHTML = `共找到 <strong>${list.length}</strong> 家门店 · 已定位 <strong>${escapeHtml(targetStore.name)}</strong>`;
    } else {
      resultCountEl.innerHTML = `共找到 <strong>${list.length}</strong> 家门店`;
    }

    const frag = document.createDocumentFragment();
    for (const store of list) {
      frag.appendChild(renderStoreItem(store));
    }
    storeListEl.appendChild(frag);

    if (scrollToId && targetStore) {
      scrollToStoreCard(scrollToId);
    }

    return true;
  }

  function showStoreNotFound(storeId) {
    provinceEl.value = "";
    fillCityOptions("");
    cityEl.value = "";
    keywordEl.value = "";
    clearKeywordEl.classList.add("hidden");
    storeListEl.innerHTML = "";
    storeListEl.classList.add("hidden");
    emptyStateEl.classList.remove("hidden");
    emptyStateEl.querySelector(".empty-title").textContent = "未找到该门店";
    emptyStateEl.querySelector(".empty-desc").textContent =
      `门店编号「${storeId}」不存在或已下线，请检查链接参数 storeId 是否正确。`;
    resultCountEl.textContent = "";
  }

  function applyStoreDeepLink(storeId) {
    const store = findStoreById(storeId);
    if (!store) {
      showStoreNotFound(storeId);
      return;
    }

    deepLinkTargetId = String(store.id);
    provinceEl.value = store.province;
    fillCityOptions(store.province);
    cityEl.value = store.city;
    keywordEl.value = "";
    clearKeywordEl.classList.add("hidden");
    renderList({ scrollToId: store.id });
  }

  function resetFilters() {
    deepLinkTargetId = null;
    provinceEl.value = "";
    fillCityOptions("");
    cityEl.value = "";
    keywordEl.value = "";
    clearKeywordEl.classList.add("hidden");
    emptyStateEl.querySelector(".empty-title").textContent = "暂无匹配门店";
    emptyStateEl.querySelector(".empty-desc").textContent =
      "试试调整省份、城市或搜索关键词";
    renderList();
  }

  provinceEl.addEventListener("change", () => {
    deepLinkTargetId = null;
    fillCityOptions(provinceEl.value);
    cityEl.value = "";
    renderList();
  });

  cityEl.addEventListener("change", () => {
    deepLinkTargetId = null;
    renderList();
  });

  keywordEl.addEventListener("input", () => {
    deepLinkTargetId = null;
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

  fillProvinceOptions();
  fillCityOptions("");

  const deepLinkStoreId = parseStoreIdFromUrl();
  if (deepLinkStoreId) {
    applyStoreDeepLink(deepLinkStoreId);
  } else {
    renderList();
  }

  function initStoresScrollUi() {
    const stickyBar = document.getElementById("storesStickyBar");
    const stickySentinel = document.getElementById("storesStickySentinel");
    const backToTopBtn = document.getElementById("backToTop");
    if (!stickyBar || !stickySentinel || !backToTopBtn) return;

    const stickyTopPx = () => {
      const topBar = document.querySelector(".top-bar");
      const h = topBar ? topBar.getBoundingClientRect().height : 44;
      return `${Math.ceil(h)}px`;
    };

    document.documentElement.style.setProperty("--stores-sticky-top", stickyTopPx());

    const stickyObserver = new IntersectionObserver(
      ([entry]) => {
        stickyBar.classList.toggle("is-compact", !entry.isIntersecting);
      },
      {
        root: null,
        threshold: 0,
        rootMargin: `-${stickyTopPx()} 0px 0px 0px`,
      }
    );
    stickyObserver.observe(stickySentinel);

    let scrollTicking = false;
    function updateBackToTop() {
      backToTopBtn.classList.toggle("is-visible", window.scrollY > 360);
      scrollTicking = false;
    }

    window.addEventListener(
      "scroll",
      () => {
        if (!scrollTicking) {
          scrollTicking = true;
          requestAnimationFrame(updateBackToTop);
        }
      },
      { passive: true }
    );
    updateBackToTop();

    backToTopBtn.addEventListener("click", () => {
      window.scrollTo({ top: 0, behavior: "smooth" });
    });

    window.addEventListener("resize", () => {
      document.documentElement.style.setProperty("--stores-sticky-top", stickyTopPx());
    });
  }

  initStoresScrollUi();
})();
