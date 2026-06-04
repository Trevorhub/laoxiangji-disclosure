(function () {
  const modal = window.LicenseModal;
  modal.init();

  const keywordEl = document.getElementById("keyword");
  const clearKeywordEl = document.getElementById("clearKeyword");
  const platformListEl = document.getElementById("platformList");
  const emptyStateEl = document.getElementById("emptyState");
  const resultCountEl = document.getElementById("resultCount");
  const resetFiltersEl = document.getElementById("resetFilters");

  const LICENSE_LABELS = {
    business: "平台营业执照",
    operating: "食品经营许可证",
  };

  const LICENSE_TYPES = ["business", "operating"];

  const DOC_ICON =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm-1 2.5L18.5 9H13V4.5zM8 13h8v1.5H8V13zm0 3.5h5V18H8v-1.5z"/></svg>';

  function getAvailableLicenseTypes(platform) {
    return LICENSE_TYPES.filter((type) => !!platform.licenses[type]);
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function getFilteredPlatforms() {
    const keyword = keywordEl.value.trim().toLowerCase();
    if (!keyword) return PLATFORMS;
    return PLATFORMS.filter((item) => {
      const haystack = `${item.name} ${item.desc} ${item.channel}`.toLowerCase();
      return haystack.includes(keyword);
    });
  }

  function buildLicenseBlocks(platform) {
    return getAvailableLicenseTypes(platform).map((type) => ({
      type,
      label: LICENSE_LABELS[type],
      src: platform.licenses[type],
      alt: `${platform.name} ${LICENSE_LABELS[type]}`,
    }));
  }

  function openAllLicensesModal(platform) {
    const blocks = buildLicenseBlocks(platform);
    if (blocks.length === 0) return;
    modal.openMultiple({
      tag: "平台资质",
      title: "平台证照公示",
      subtitle: platform.name,
      blocks,
    });
  }

  function openLicenseModal(platform, type) {
    if (!platform.licenses[type]) return;
    const label = LICENSE_LABELS[type];
    modal.openSingle({
      tag: "平台资质",
      title: label,
      subtitle: platform.name,
      src: platform.licenses[type],
      alt: `${platform.name} ${label}`,
    });
  }

  function renderLicenseChips(platform) {
    return getAvailableLicenseTypes(platform).map(
      (type) => `
        <button type="button" class="license-chip" data-license="${type}">
          ${DOC_ICON}${LICENSE_LABELS[type]}
        </button>
      `
    ).join("");
  }

  function renderPlatformItem(platform) {
    const li = document.createElement("li");
    li.className = "store-card";
    li.innerHTML = `
      <div class="store-card__body" role="button" tabindex="0" aria-label="查看${escapeHtml(platform.name)}全部证照">
        <div class="store-card__head">
          <span class="city-tag">${escapeHtml(platform.channel)}</span>
        </div>
        <h3 class="store-name">${escapeHtml(platform.name)}</h3>
        <p class="store-desc">${escapeHtml(platform.desc)}</p>
      </div>
      <div class="store-licenses store-licenses--platform">
        ${renderLicenseChips(platform)}
      </div>
    `;

    const cardBody = li.querySelector(".store-card__body");
    cardBody.addEventListener("click", () => openAllLicensesModal(platform));
    cardBody.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openAllLicensesModal(platform);
      }
    });

    li.querySelectorAll(".license-chip").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        openLicenseModal(platform, btn.dataset.license);
      });
    });

    return li;
  }

  function renderList() {
    const list = getFilteredPlatforms();
    platformListEl.innerHTML = "";

    if (list.length === 0) {
      emptyStateEl.classList.remove("hidden");
      platformListEl.classList.add("hidden");
      resultCountEl.textContent = "";
      return;
    }

    emptyStateEl.classList.add("hidden");
    platformListEl.classList.remove("hidden");
    resultCountEl.innerHTML = `共 <strong>${list.length}</strong> 项平台`;

    const frag = document.createDocumentFragment();
    for (const platform of list) {
      frag.appendChild(renderPlatformItem(platform));
    }
    platformListEl.appendChild(frag);
  }

  function resetFilters() {
    keywordEl.value = "";
    clearKeywordEl.classList.add("hidden");
    renderList();
  }

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

  renderList();
})();
