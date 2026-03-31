(() => {
  const PAGE_SIZE = 30;
  const state = {
    payload: null,
    filteredItems: [],
    visibleCount: PAGE_SIZE,
    activeType: "",
    sortMode: "newest",
    specialFilter: "",
    activePreset: "recent",
  };

  const BROKEN_TEXT_REPLACEMENTS = [
    ["N┬║", "Nº"],
    ["MAR├cO", "MARÇO"],
    ["D├u", "DÁ"],
    ["PROVID├eNCIAS", "PROVIDÊNCIAS"],
    ["JOS├e", "JOSÉ"],
    ["J├UNIOR", "JÚNIOR"],
    ["S├uo", "São"],
    ["atribui├º├Aes", "atribuições"],
    ["Org├onica", "Orgânica"],
    ["Munic├¡pio", "Município"],
  ];

  const elements = {
    headlineMetrics: document.getElementById("headline-metrics"),
    heroUpdatedAt: document.getElementById("hero-updated-at"),
    heroStatusCopy: document.getElementById("hero-status-copy"),
    heroDiaryCount: document.getElementById("hero-diary-count"),
    heroSupplierCount: document.getElementById("hero-supplier-count"),
    executiveMeta: document.getElementById("executive-meta"),
    executiveOverview: document.getElementById("executive-overview"),
    executiveCards: document.getElementById("executive-cards"),
    prioritiesMeta: document.getElementById("priorities-meta"),
    priorityCards: document.getElementById("priority-cards"),
    quickFilters: document.getElementById("quick-filters"),
    activeFilters: document.getElementById("active-filters"),
    searchInput: document.getElementById("search-input"),
    typeSelect: document.getElementById("type-select"),
    yearSelect: document.getElementById("year-select"),
    typePills: document.getElementById("type-pills"),
    resetFilters: document.getElementById("reset-filters"),
    resultsMeta: document.getElementById("results-meta"),
    contractsList: document.getElementById("contracts-list"),
    loadMore: document.getElementById("load-more"),
    footerCopy: document.getElementById("footer-copy"),
  };

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function repairBrokenText(value) {
    let text = String(value ?? "");
    BROKEN_TEXT_REPLACEMENTS.forEach(([broken, fixed]) => {
      text = text.split(broken).join(fixed);
    });
    return text;
  }

  function normalizePayload(value) {
    if (typeof value === "string") {
      return repairBrokenText(value);
    }

    if (Array.isArray(value)) {
      return value.map((item) => normalizePayload(item));
    }

    if (value && typeof value === "object") {
      return Object.fromEntries(
        Object.entries(value).map(([key, entryValue]) => [key, normalizePayload(entryValue)]),
      );
    }

    return value;
  }

  function normalizeForSearch(value) {
    return repairBrokenText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function formatCurrency(value) {
    return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(Number(value || 0));
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("pt-BR").format(Number(value || 0));
  }

  function formatPercent(value) {
    return new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 1 }).format(Number(value || 0));
  }

  function formatDate(value, includeTime = false) {
    if (!value) return "Sem data";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return new Intl.DateTimeFormat("pt-BR", includeTime ? { dateStyle: "medium", timeStyle: "short" } : { dateStyle: "medium" }).format(date);
  }

  function truncateText(value, limit = 88) {
    const text = String(value || "").replace(/\s+/g, " ").trim();
    if (text.length <= limit) return text;
    return `${text.slice(0, limit - 3).trim()}...`;
  }

  function getGovernmentLabel(sphere) {
    return {
      municipal: "Governo Municipal",
      estadual: "Governo Estadual",
      federal: "Governo Federal",
    }[String(sphere || "").toLowerCase()] || "";
  }

  function formatOrganizationLabel(name, sphere) {
    const normalizedName = String(name || "").trim();
    if (!normalizedName) return "Não identificado";

    const governmentLabel = getGovernmentLabel(sphere);
    return governmentLabel ? `${normalizedName} - ${governmentLabel}` : normalizedName;
  }

  function getItems() {
    return state.payload?.items || [];
  }

  function sortByNewest(items) {
    return [...items].sort((a, b) => new Date(b.publishedAt || 0) - new Date(a.publishedAt || 0) || Number(b.edition || 0) - Number(a.edition || 0));
  }

  function sortByHighestValue(items) {
    return [...items].sort((a, b) => Number(b.valueNumber || 0) - Number(a.valueNumber || 0) || new Date(b.publishedAt || 0) - new Date(a.publishedAt || 0));
  }

  function getLatestYearEntry() {
    return state.payload?.yearSummary?.[0] || null;
  }

  function getTopTypeEntry() {
    return state.payload?.typeSummary?.[0] || null;
  }

  function populateSelect(select, entries, emptyLabel) {
    select.innerHTML = "";
    [{ value: "", label: emptyLabel }, ...entries].forEach((entry) => {
      const option = document.createElement("option");
      option.value = entry.value;
      option.textContent = entry.label;
      select.appendChild(option);
    });
  }

  function renderMetrics() {
    const summary = state.payload?.summary || {};
    const cards = [
      { label: "Registros", value: formatNumber(summary.totalItems || 0), meta: "atos organizados" },
      { label: "Valor identificado", value: formatCurrency(summary.totalValue || 0), meta: "somatório extraído" },
      { label: "Alta confiança", value: formatNumber(summary.highConfidenceItems || 0), meta: "itens consistentes" },
    ];

    elements.headlineMetrics.innerHTML = cards.map((card) => `
      <article class="metric-card">
        <strong>${escapeHtml(card.value)}</strong>
        <span>${escapeHtml(card.label)}</span>
        <span>${escapeHtml(card.meta)}</span>
      </article>
    `).join("");

    elements.heroUpdatedAt.textContent = `Atualizado em ${formatDate(state.payload?.generatedAt, true)}`;
    elements.heroStatusCopy.textContent = `${formatNumber(summary.totalItems || 0)} registros públicos organizados para consulta, com leitura consolidada do Diário Oficial.`;
    elements.heroDiaryCount.textContent = formatNumber(summary.analyzedDiaryCount || 0);
    elements.heroSupplierCount.textContent = formatNumber(summary.uniqueSuppliers || 0);
    elements.footerCopy.textContent = `Atualizado em ${formatDate(state.payload?.generatedAt, true)}. ${formatNumber(summary.analyzedDiaryCount || 0)} edições analisadas e ${formatNumber(summary.uniqueSuppliers || 0)} fornecedores identificados.`;
  }

  function renderExecutiveSummary() {
    const summary = state.payload?.summary || {};
    const topType = state.payload?.typeSummary?.[0];
    const topOrganization = state.payload?.organizationSummary?.[0];
    const years = state.payload?.yearSummary || [];
    const newestItem = sortByNewest(getItems())[0];
    const highConfidenceRate = summary.totalItems ? summary.highConfidenceItems / summary.totalItems : 0;
    const coverageStart = years.length ? years[years.length - 1].year : "Sem recorte";
    const coverageEnd = years.length ? years[0].year : "Sem recorte";

    elements.executiveMeta.textContent = `Base consolidada entre ${coverageStart} e ${coverageEnd}, com última publicação em ${formatDate(newestItem?.publishedAt)}.`;
    elements.executiveOverview.innerHTML = `
      <article class="executive-overview-card">
        <span class="executive-overview-label">Síntese executiva</span>
        <h3>${escapeHtml(formatNumber(summary.totalItems || 0))} registros públicos reunidos para leitura gerencial</h3>
        <p>
          A base atual combina ${escapeHtml(formatNumber(summary.analyzedDiaryCount || 0))} edições analisadas,
          ${escapeHtml(formatNumber(summary.uniqueSuppliers || 0))} fornecedores identificados e valor conhecido de
          ${escapeHtml(formatCurrency(summary.totalValue || 0))}.
        </p>
      </article>
    `;

    const cards = [
      {
        label: "Confiabilidade da leitura",
        value: formatPercent(highConfidenceRate),
        meta: `${formatNumber(summary.highConfidenceItems || 0)} registros com alta confiança.`,
      },
      {
        label: "Tipo predominante",
        value: topType?.type || "Sem classificação",
        meta: `${formatNumber(topType?.count || 0)} registros no principal agrupamento.`,
      },
      {
        label: "Maior concentração institucional",
        value: topOrganization?.displayName || "Não identificado",
        meta: `${formatNumber(topOrganization?.count || 0)} registros no órgão com maior volume.`,
      },
      {
        label: "Recência da base",
        value: formatDate(newestItem?.publishedAt),
        meta: "Data do registro público mais recente presente na base.",
      },
    ];

    elements.executiveCards.innerHTML = cards.map((card) => `
      <article class="executive-card">
        <span>${escapeHtml(card.label)}</span>
        <strong>${escapeHtml(card.value)}</strong>
        <p>${escapeHtml(card.meta)}</p>
      </article>
    `).join("");
  }

  function renderPriorities() {
    const items = getItems();
    const latestItems = sortByNewest(items).slice(0, 3);
    const highestValueItems = [...items]
      .filter((item) => Number(item.valueNumber || 0) > 0)
      .sort((a, b) => Number(b.valueNumber || 0) - Number(a.valueNumber || 0))
      .slice(0, 3);
    const itemsWithoutValue = sortByNewest(items.filter((item) => Number(item.valueNumber || 0) <= 0));
    const missingValueSample = itemsWithoutValue.slice(0, 3);
    const topOrganizations = (state.payload?.organizationSummary || []).slice(0, 3);

    elements.prioritiesMeta.textContent = "4 frentes organizadas para leitura inicial da chefia: recência, maior valor, lacunas de valor e concentração institucional.";

    const cards = [
      {
        label: "Mais recentes",
        tone: "recent",
        headline: latestItems[0] ? formatDate(latestItems[0].publishedAt) : "Sem data recente",
        summary: "Últimos atos inseridos na base publicada.",
        items: latestItems.map((item) => ({
          primary: truncateText(item.title || "Ato contratual"),
          secondary: `${formatDate(item.publishedAt)} | ${truncateText(item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere), 52)}`,
        })),
      },
      {
        label: "Maiores valores",
        tone: "value",
        headline: highestValueItems[0]?.value || "Sem valor identificado",
        summary: "Registros com maior impacto financeiro identificado na base.",
        items: highestValueItems.map((item) => ({
          primary: truncateText(item.title || "Ato contratual"),
          secondary: `${item.value || "Valor não informado"} | ${truncateText(item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere), 52)}`,
        })),
      },
      {
        label: "Sem valor informado",
        tone: "attention",
        headline: formatNumber(itemsWithoutValue.length),
        summary: "Registros ainda sem valor consolidado para leitura rápida.",
        items: missingValueSample.map((item) => ({
          primary: truncateText(item.title || "Ato contratual"),
          secondary: `${formatDate(item.publishedAt)} | ${truncateText(item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere), 52)}`,
        })),
      },
      {
        label: "Maior concentração por órgão",
        tone: "institution",
        headline: topOrganizations[0]?.displayName || "Não identificado",
        summary: "Órgãos com maior volume de registros na base publicada.",
        items: topOrganizations.map((item) => ({
          primary: truncateText(item.displayName || "Órgão não identificado", 72),
          secondary: `${formatNumber(item.count || 0)} registro(s) | ${formatCurrency(item.totalValue || 0)}`,
        })),
      },
    ];

    elements.priorityCards.innerHTML = cards.map((card) => `
      <article class="priority-card priority-card-${escapeHtml(card.tone)}">
        <div class="priority-card-head">
          <span class="priority-label">${escapeHtml(card.label)}</span>
          <strong>${escapeHtml(card.headline)}</strong>
          <p>${escapeHtml(card.summary)}</p>
        </div>
        <div class="priority-list">
          ${card.items.map((item) => `
            <article class="priority-list-item">
              <strong>${escapeHtml(item.primary)}</strong>
              <span>${escapeHtml(item.secondary)}</span>
            </article>
          `).join("")}
        </div>
      </article>
    `).join("");
  }

  function renderQuickFilters() {
    const latestYear = getLatestYearEntry();
    const topType = getTopTypeEntry();
    const highestValueItem = [...getItems()]
      .filter((item) => Number(item.valueNumber || 0) > 0)
      .sort((a, b) => Number(b.valueNumber || 0) - Number(a.valueNumber || 0))[0];
    const missingValueCount = getItems().filter((item) => Number(item.valueNumber || 0) <= 0).length;

    const presets = [
      {
        id: "recent",
        label: "Mais recentes",
        meta: "Ordem cronológica padrão",
      },
      {
        id: "latestYear",
        label: latestYear ? `Ano ${latestYear.year}` : "Ano mais recente",
        meta: latestYear ? `${formatNumber(latestYear.count)} registro(s)` : "Sem recorte anual",
      },
      {
        id: "topType",
        label: "Tipo predominante",
        meta: topType ? `${topType.type} (${formatNumber(topType.count)})` : "Sem classificação",
      },
      {
        id: "highestValue",
        label: "Maiores valores",
        meta: highestValueItem ? `Até ${highestValueItem.value || formatCurrency(highestValueItem.valueNumber)}` : "Sem valor identificado",
      },
      {
        id: "missingValue",
        label: "Sem valor",
        meta: `${formatNumber(missingValueCount)} registro(s)`,
      },
    ];

    elements.quickFilters.innerHTML = presets.map((preset) => `
      <button class="quick-filter-button ${state.activePreset === preset.id ? "active" : ""}" data-preset="${escapeHtml(preset.id)}" type="button">
        <span>${escapeHtml(preset.label)}</span>
        <strong>${escapeHtml(preset.meta)}</strong>
      </button>
    `).join("");

    [...elements.quickFilters.querySelectorAll("[data-preset]")].forEach((button) => {
      button.addEventListener("click", () => {
        applyQuickPreset(button.dataset.preset);
      });
    });
  }

  function renderActiveFiltersSummary({ query, type, year }) {
    const parts = [];
    const hasManualFilter = Boolean(query || type || year || state.specialFilter === "missingValue" || state.sortMode === "highestValue" || state.activePreset && state.activePreset !== "recent");

    if (!hasManualFilter) {
      elements.activeFilters.textContent = "Sem filtros ativos. A lista abaixo segue do mais novo para o mais antigo.";
      return;
    }

    if (state.activePreset === "recent") {
      parts.push("Atalho: mais recentes");
    }
    if (query) {
      parts.push(`Busca: ${query}`);
    }
    if (type) {
      parts.push(`Tipo: ${type}`);
    }
    if (year) {
      parts.push(`Ano: ${year}`);
    }
    if (state.specialFilter === "missingValue") {
      parts.push("Somente registros sem valor informado");
    }
    if (state.sortMode === "highestValue") {
      parts.push("Ordenação: maior valor");
    }
    else {
      parts.push("Ordenação: mais recentes");
    }

    elements.activeFilters.textContent = parts.join(" | ");
  }

  function applyQuickPreset(presetId) {
    const latestYear = getLatestYearEntry();
    const topType = getTopTypeEntry();

    elements.searchInput.value = "";
    elements.typeSelect.value = "";
    elements.yearSelect.value = "";
    state.sortMode = "newest";
    state.specialFilter = "";
    state.activePreset = presetId;

    if (presetId === "latestYear" && latestYear?.year) {
      elements.yearSelect.value = String(latestYear.year);
    }
    else if (presetId === "topType" && topType?.type) {
      elements.typeSelect.value = String(topType.type);
    }
    else if (presetId === "highestValue") {
      state.sortMode = "highestValue";
    }
    else if (presetId === "missingValue") {
      state.specialFilter = "missingValue";
    }

    updateFilters({ resetVisible: true });
  }

  function renderTypePills() {
    const rows = state.payload?.typeSummary || [];
    elements.typePills.innerHTML = rows.map((row) => `<button class="type-pill ${state.activeType === row.type ? "active" : ""}" data-type="${escapeHtml(row.type)}" type="button">${escapeHtml(row.type)} (${escapeHtml(row.count)})</button>`).join("");
    [...elements.typePills.querySelectorAll("[data-type]")].forEach((button) => {
      button.addEventListener("click", () => {
        state.activeType = state.activeType === button.dataset.type ? "" : button.dataset.type;
        elements.typeSelect.value = state.activeType;
        state.activePreset = "";
        updateFilters({ resetVisible: true });
      });
    });
  }

  function renderContracts() {
    const visibleRows = state.filteredItems.slice(0, state.visibleCount);
    const sortCopy = state.sortMode === "highestValue" ? "ordenados por maior valor" : "ordenados do mais novo para o mais antigo";
    elements.resultsMeta.textContent = `${visibleRows.length} de ${state.filteredItems.length} registro(s) exibidos, ${sortCopy}.`;
    if (!visibleRows.length) {
      elements.contractsList.innerHTML = `<div class="empty-card">Nenhum contrato corresponde aos filtros atuais.</div>`;
      elements.loadMore.hidden = true;
      return;
    }

    elements.contractsList.innerHTML = visibleRows.map((item) => `
      <article class="contract-card">
        <div class="contract-head">
          <div>
            <div class="badge-row">
              <span class="badge">${escapeHtml(item.type || "Ato")}</span>
              <span class="badge neutral">${escapeHtml(item.recordClass || "Registro")}</span>
            </div>
            <h3>${escapeHtml(item.title || "Ato contratual")}</h3>
          </div>
          <div class="meta-row">
            <span class="meta-chip">${escapeHtml(`Edição ${item.edition || "-"}`)}</span>
            <span class="meta-chip">${escapeHtml(formatDate(item.publishedAt))}</span>
          </div>
        </div>
        <div class="meta-line"><strong>Órgão:</strong> ${escapeHtml(item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere))}</div>
        <div class="meta-line"><strong>Fornecedor:</strong> ${escapeHtml(item.contractor || "Não identificado")}</div>
        <div class="meta-line"><strong>Valor:</strong> ${escapeHtml(item.value || "Não informado")}</div>
        <p class="contract-summary">${escapeHtml(item.summary || "Sem resumo consolidado.")}</p>
        <div class="action-row">
          ${item.viewUrl ? `<a class="action-link" href="${escapeHtml(item.viewUrl)}" target="_blank" rel="noopener noreferrer">Abrir edição oficial</a>` : ""}
        </div>
      </article>
    `).join("");
    elements.loadMore.hidden = visibleRows.length >= state.filteredItems.length;
  }

  function updateFilters({ resetVisible = false } = {}) {
    const query = normalizeForSearch(elements.searchInput.value.trim());
    const type = elements.typeSelect.value;
    const year = elements.yearSelect.value;
    state.activeType = type;

    const filtered = getItems().filter((item) => {
      const organizationLabel = item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere);
      const haystack = normalizeForSearch([item.title, item.organization, organizationLabel, item.contractor, item.summary, item.type, item.edition].join(" "));
      if (query && !haystack.includes(query)) return false;
      if (type && item.type !== type) return false;
      if (year && !(item.publishedAt || "").startsWith(year)) return false;
      if (state.specialFilter === "missingValue" && Number(item.valueNumber || 0) > 0) return false;
      return true;
    });

    state.filteredItems = state.sortMode === "highestValue" ? sortByHighestValue(filtered) : sortByNewest(filtered);
    if (resetVisible) state.visibleCount = PAGE_SIZE;
    renderQuickFilters();
    renderActiveFiltersSummary({ query, type, year });
    renderTypePills();
    renderContracts();
  }

  async function loadDashboard() {
    const response = await fetch("./data/dashboard.json", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Falha ao carregar a base pública: ${response.status}`);
    }

    state.payload = normalizePayload(await response.json());
    const typeOptions = (state.payload.typeSummary || []).map((item) => ({ value: item.type, label: item.type }));
    const yearOptions = (state.payload.yearSummary || []).map((item) => ({ value: item.year, label: `${item.year} (${item.count})` }));
    populateSelect(elements.typeSelect, typeOptions, "Todos os tipos");
    populateSelect(elements.yearSelect, yearOptions, "Todos os anos");
    renderMetrics();
    renderExecutiveSummary();
    renderPriorities();
    renderQuickFilters();
    updateFilters({ resetVisible: true });
  }

  function wireEvents() {
    elements.searchInput.addEventListener("input", () => {
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.typeSelect.addEventListener("change", () => {
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.yearSelect.addEventListener("change", () => {
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.resetFilters.addEventListener("click", () => {
      elements.searchInput.value = "";
      elements.typeSelect.value = "";
      elements.yearSelect.value = "";
      state.activeType = "";
      state.sortMode = "newest";
      state.specialFilter = "";
      state.activePreset = "recent";
      updateFilters({ resetVisible: true });
    });
    elements.loadMore.addEventListener("click", () => {
      state.visibleCount += PAGE_SIZE;
      renderContracts();
    });
  }

  wireEvents();
  loadDashboard().catch((error) => {
    elements.headlineMetrics.innerHTML = `<div class="empty-card">${escapeHtml(error.message)}</div>`;
    elements.executiveOverview.innerHTML = `<div class="empty-card">${escapeHtml(error.message)}</div>`;
    elements.executiveCards.innerHTML = "";
    elements.executiveMeta.textContent = "Não foi possível consolidar o panorama gerencial.";
    elements.priorityCards.innerHTML = "";
    elements.prioritiesMeta.textContent = "Não foi possível consolidar os pontos de atenção.";
    elements.quickFilters.innerHTML = "";
    elements.activeFilters.textContent = "Não foi possível consolidar os atalhos de navegação.";
    elements.contractsList.innerHTML = `<div class="empty-card">${escapeHtml(error.message)}</div>`;
    elements.resultsMeta.textContent = "Não foi possível carregar a base pública.";
  });
})();
