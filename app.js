(() => {
  const PAGE_SIZE = 30;
  const FEEDBACK_TIMEOUT_MS = 2400;
  const DEFAULT_VIEW = "overview";
  const VALID_VIEWS = new Set(["overview", "consulta", "contratos"]);
  const SORT_LABELS = {
    newest: "mais recentes",
    highestValue: "maior valor",
    oldest: "mais antigos",
    alphabetical: "ordem alfabética",
  };
  const CONFIDENCE_ORDER = ["alta", "media", "média", "baixa"];
  const state = {
    payload: null,
    filteredItems: [],
    visibleCount: PAGE_SIZE,
    activeType: "",
    sortMode: "newest",
    specialFilter: "",
    activePreset: "",
    currentView: DEFAULT_VIEW,
    feedbackTimer: 0,
  };
  const initialUrlState = readUrlState();
  const BROKEN_TEXT_REPLACEMENTS = [
    ["N┬║", "Nº"],
    ["MAR├cO", "MARÇO"],
    ["D├u", "DÁ"],
    ["PROVID├eNCIAS", "PROVIDÊNCIAS"],
    ["JOS├e", "JOSÉ"],
    ["J├UNIOR", "JÚNIOR"],
    ["S├uo", "São"],
    ["s├uo", "são"],
    ["atribui├º├Aes", "atribuições"],
    ["Org├onica", "Orgânica"],
    ["Munic├¡pio", "Município"],
    ["Nâ”¬â•‘", "Nº"],
    ["MARâ”œcO", "MARÇO"],
    ["Dâ”œu", "DÁ"],
    ["PROVIDâ”œeNCIAS", "PROVIDÊNCIAS"],
    ["JOSâ”œe", "JOSÉ"],
    ["Jâ”œUNIOR", "JÚNIOR"],
    ["Sâ”œuo", "São"],
    ["sâ”œuo", "são"],
    ["atribuiâ”œÂºâ”œAes", "atribuições"],
    ["Orgâ”œonica", "Orgânica"],
    ["Municâ”œÂ¡pio", "Município"],
  ];

  const elements = {
    headlineMetrics: document.getElementById("headline-metrics"),
    heroSearchForm: document.getElementById("hero-search-form"),
    heroSearchInput: document.getElementById("hero-search-input"),
    heroUpdatedAt: document.getElementById("hero-updated-at"),
    heroStatusCopy: document.getElementById("hero-status-copy"),
    heroDiaryCount: document.getElementById("hero-diary-count"),
    heroSupplierCount: document.getElementById("hero-supplier-count"),
    executiveMeta: document.getElementById("executive-meta"),
    executiveOverview: document.getElementById("executive-overview"),
    executiveCards: document.getElementById("executive-cards"),
    prioritiesMeta: document.getElementById("priorities-meta"),
    priorityCards: document.getElementById("priority-cards"),
    insightsMeta: document.getElementById("insights-meta"),
    insightCards: document.getElementById("insight-cards"),
    quickFilters: document.getElementById("quick-filters"),
    activeFilters: document.getElementById("active-filters"),
    shareFeedback: document.getElementById("share-feedback"),
    selectionSummary: document.getElementById("selection-summary"),
    searchInput: document.getElementById("search-input"),
    typeSelect: document.getElementById("type-select"),
    yearSelect: document.getElementById("year-select"),
    administrationSelect: document.getElementById("administration-select"),
    confidenceSelect: document.getElementById("confidence-select"),
    sortSelect: document.getElementById("sort-select"),
    typePills: document.getElementById("type-pills"),
    copyLink: document.getElementById("copy-link"),
    exportCsv: document.getElementById("export-csv"),
    resetFilters: document.getElementById("reset-filters"),
    resultsMeta: document.getElementById("results-meta"),
    contractsList: document.getElementById("contracts-list"),
    loadMore: document.getElementById("load-more"),
    footerCopy: document.getElementById("footer-copy"),
    viewButtons: [...document.querySelectorAll("[data-view-button]")],
    viewLinks: [...document.querySelectorAll("[data-view-target]")],
    pageViews: [...document.querySelectorAll("[data-page-view]")],
  };

  function readUrlState() {
    const params = new URLSearchParams(window.location.search);
    const sort = params.get("sort");
    return {
      query: params.get("q") || "",
      type: params.get("type") || "",
      year: params.get("year") || "",
      administration: params.get("administration") || "",
      confidence: params.get("confidence") || "",
      sort: SORT_LABELS[sort] ? sort : "newest",
      special: params.get("special") === "missingValue" ? "missingValue" : "",
      preset: params.get("preset") || "",
      view: VALID_VIEWS.has(params.get("view")) ? params.get("view") : DEFAULT_VIEW,
    };
  }

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
    if (typeof value === "string") return repairBrokenText(value);
    if (Array.isArray(value)) return value.map((item) => normalizePayload(item));
    if (value && typeof value === "object") {
      return Object.fromEntries(Object.entries(value).map(([key, entryValue]) => [key, normalizePayload(entryValue)]));
    }
    return value;
  }

  function normalizeForSearch(value) {
    return repairBrokenText(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
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
    const text = repairBrokenText(String(value || "").replace(/\s+/g, " ").trim());
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
    const normalizedName = repairBrokenText(String(name || "").trim());
    if (!normalizedName) return "Não identificado";
    const governmentLabel = getGovernmentLabel(sphere);
    return governmentLabel ? `${normalizedName} - ${governmentLabel}` : normalizedName;
  }

  function formatConfidenceLabel(value) {
    return {
      alta: "Alta confiança",
      media: "Leitura média",
      "média": "Leitura média",
      baixa: "Baixa confiança",
    }[String(value || "").toLowerCase()] || "Sem classificação";
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

  function sortByOldest(items) {
    return [...items].sort((a, b) => new Date(a.publishedAt || 0) - new Date(b.publishedAt || 0) || Number(a.edition || 0) - Number(b.edition || 0));
  }

  function sortAlphabetically(items) {
    return [...items].sort((a, b) => repairBrokenText(a.title || "").localeCompare(repairBrokenText(b.title || ""), "pt-BR") || new Date(b.publishedAt || 0) - new Date(a.publishedAt || 0));
  }

  function sortItems(items) {
    if (state.sortMode === "highestValue") return sortByHighestValue(items);
    if (state.sortMode === "oldest") return sortByOldest(items);
    if (state.sortMode === "alphabetical") return sortAlphabetically(items);
    return sortByNewest(items);
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

  function setSelectIfValid(select, value) {
    if (!value) {
      select.value = "";
      return;
    }
    select.value = [...select.options].some((option) => option.value === value) ? value : "";
  }

  function getLatestYearEntry() {
    return state.payload?.yearSummary?.[0] || null;
  }

  function getTopTypeEntry() {
    return state.payload?.typeSummary?.[0] || null;
  }

  function getAdministrationValueForDate(value) {
    if (!value) return "";
    const date = new Date(value);
    const yearNumber = Number.isNaN(date.getTime())
      ? Number(String(value).slice(0, 4))
      : date.getFullYear();
    if (!Number.isFinite(yearNumber) || yearNumber <= 0) return "";
    const start = yearNumber - ((yearNumber - 1) % 4);
    return `${start}-${start + 3}`;
  }

  function formatAdministrationLabel(value) {
    return value ? `Gestão ${value}` : "";
  }

  function buildAdministrationEntries(items) {
    const counts = new Map();
    items.forEach((item) => {
      const administration = getAdministrationValueForDate(item.publishedAt);
      if (!administration) return;
      counts.set(administration, (counts.get(administration) || 0) + 1);
    });
    return [...counts.entries()]
      .sort((a, b) => Number(b[0].split("-")[0]) - Number(a[0].split("-")[0]))
      .map(([value, count]) => ({ value, label: `${formatAdministrationLabel(value)} (${formatNumber(count)})` }));
  }

  function administrationIncludesYear(administration, year) {
    if (!administration || !year) return true;
    const yearNumber = Number(String(year).slice(0, 4));
    const [start, end] = String(administration).split("-").map(Number);
    if (!Number.isFinite(yearNumber) || !Number.isFinite(start) || !Number.isFinite(end)) return false;
    return yearNumber >= start && yearNumber <= end;
  }

  function applyView(view, { scroll = false, sync = true } = {}) {
    const nextView = VALID_VIEWS.has(view) ? view : DEFAULT_VIEW;
    state.currentView = nextView;

    elements.pageViews.forEach((pageView) => {
      const isActive = pageView.dataset.pageView === nextView;
      pageView.hidden = !isActive;
      pageView.classList.toggle("is-active", isActive);
    });

    elements.viewButtons.forEach((button) => {
      const isActive = button.dataset.viewButton === nextView;
      button.classList.toggle("active", isActive);
      button.setAttribute("aria-pressed", String(isActive));
    });

    if (sync) syncUrlState();
    if (scroll) {
      document.getElementById(`page-view-${nextView}`)?.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  function getPresetLabel(presetId) {
    return {
      recent: "mais recentes",
      latestYear: "ano mais recente",
      topType: "tipo predominante",
      highestValue: "maiores valores",
      highConfidence: "alta confiança",
      missingValue: "sem valor informado",
    }[presetId] || presetId;
  }

  function buildConfidenceEntries(items) {
    const counts = new Map();
    items.forEach((item) => {
      const key = String(item.confidence || "").toLowerCase();
      if (!key) return;
      counts.set(key, (counts.get(key) || 0) + 1);
    });
    return [...counts.entries()]
      .sort((a, b) => (CONFIDENCE_ORDER.indexOf(a[0]) === -1 ? 99 : CONFIDENCE_ORDER.indexOf(a[0])) - (CONFIDENCE_ORDER.indexOf(b[0]) === -1 ? 99 : CONFIDENCE_ORDER.indexOf(b[0])))
      .map(([value, count]) => ({ value, label: `${formatConfidenceLabel(value)} (${formatNumber(count)})` }));
  }

  function buildSupplierRanking(items) {
    const suppliers = new Map();
    items.forEach((item) => {
      const contractor = repairBrokenText(String(item.contractor || "").trim());
      if (!contractor) return;
      const current = suppliers.get(contractor) || { name: contractor, count: 0, totalValue: 0 };
      current.count += 1;
      current.totalValue += Number(item.valueNumber || 0);
      suppliers.set(contractor, current);
    });
    return [...suppliers.values()].sort((a, b) => b.count - a.count || b.totalValue - a.totalValue || a.name.localeCompare(b.name, "pt-BR"));
  }

  function computeSelectionSummary(items) {
    return items.reduce((summary, item) => {
      const valueNumber = Number(item.valueNumber || 0);
      if (valueNumber > 0) {
        summary.totalValue += valueNumber;
        summary.withValue += 1;
      }
      if (String(item.contractor || "").trim()) summary.withSupplier += 1;
      if (String(item.confidence || "").toLowerCase() === "alta") summary.highConfidence += 1;
      return summary;
    }, { totalValue: 0, withValue: 0, withSupplier: 0, highConfidence: 0 });
  }

  function computeLatestWave(items) {
    const newestItem = sortByNewest(items)[0];
    if (!newestItem) return null;
    const waveItems = items.filter((item) => {
      const sameDiary = newestItem.diaryId && item.diaryId && item.diaryId === newestItem.diaryId;
      const sameEdition = !newestItem.diaryId && item.edition === newestItem.edition && item.publishedAt === newestItem.publishedAt;
      return sameDiary || sameEdition;
    });
    return { edition: newestItem.edition || "-", publishedAt: newestItem.publishedAt, count: waveItems.length || 1 };
  }

  function setFeedback(message, tone = "info") {
    if (!elements.shareFeedback) return;
    window.clearTimeout(state.feedbackTimer);
    elements.shareFeedback.textContent = message;
    elements.shareFeedback.className = `inline-feedback is-${tone}`;
    state.feedbackTimer = window.setTimeout(() => {
      elements.shareFeedback.textContent = "";
      elements.shareFeedback.className = "inline-feedback";
    }, FEEDBACK_TIMEOUT_MS);
  }

  async function copyText(value) {
    try {
      if (navigator.clipboard?.writeText) {
        await navigator.clipboard.writeText(value);
        return true;
      }
    }
    catch {}

    const helper = document.createElement("textarea");
    helper.value = value;
    helper.setAttribute("readonly", "");
    helper.style.position = "absolute";
    helper.style.left = "-9999px";
    document.body.appendChild(helper);
    helper.select();
    const copied = document.execCommand("copy");
    document.body.removeChild(helper);
    return copied;
  }

  function toCsvValue(value) {
    return `"${repairBrokenText(String(value ?? "")).replace(/"/g, "\"\"")}"`;
  }

  function syncUrlState() {
    const params = new URLSearchParams();
    const query = elements.searchInput.value.trim();
    const type = elements.typeSelect.value;
    const year = elements.yearSelect.value;
    const administration = elements.administrationSelect.value;
    const confidence = elements.confidenceSelect.value;

    if (query) params.set("q", query);
    if (type) params.set("type", type);
    if (year) params.set("year", year);
    if (administration) params.set("administration", administration);
    if (confidence) params.set("confidence", confidence);
    if (state.sortMode !== "newest") params.set("sort", state.sortMode);
    if (state.specialFilter) params.set("special", state.specialFilter);
    if (state.activePreset) params.set("preset", state.activePreset);
    if (state.currentView !== DEFAULT_VIEW) params.set("view", state.currentView);

    const queryString = params.toString();
    window.history.replaceState(null, "", `${window.location.pathname}${queryString ? `?${queryString}` : ""}`);
  }

  function applyInitialUrlState() {
    elements.searchInput.value = repairBrokenText(initialUrlState.query);
    if (elements.heroSearchInput) {
      elements.heroSearchInput.value = repairBrokenText(initialUrlState.query);
    }
    setSelectIfValid(elements.typeSelect, repairBrokenText(initialUrlState.type));
    setSelectIfValid(elements.yearSelect, initialUrlState.year);
    setSelectIfValid(elements.administrationSelect, initialUrlState.administration);
    setSelectIfValid(elements.confidenceSelect, initialUrlState.confidence);
    setSelectIfValid(elements.sortSelect, initialUrlState.sort);
    state.sortMode = elements.sortSelect.value || "newest";
    state.specialFilter = initialUrlState.special;
    state.activePreset = initialUrlState.preset;
    state.currentView = initialUrlState.view;
  }

  function renderMetrics() {
    const summary = state.payload?.summary || {};
    const cards = [
      { tone: "records", kicker: "Volume", label: "Registros", value: formatNumber(summary.totalItems || 0), meta: "atos organizados" },
      { tone: "value", kicker: "Montante", label: "Valor identificado", value: formatCurrency(summary.totalValue || 0), meta: "somatório extraído" },
      { tone: "trust", kicker: "Qualidade", label: "Alta confiança", value: formatNumber(summary.highConfidenceItems || 0), meta: "itens consistentes" },
    ];

    elements.headlineMetrics.innerHTML = cards.map((card) => `
      <article class="metric-card metric-card-${escapeHtml(card.tone)}">
        <span class="metric-card-kicker">${escapeHtml(card.kicker)}</span>
        <strong>${escapeHtml(card.value)}</strong>
        <span class="metric-card-label">${escapeHtml(card.label)}</span>
        <span class="metric-card-meta">${escapeHtml(card.meta)}</span>
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
      { tone: "trust", label: "Confiabilidade da leitura", value: formatPercent(highConfidenceRate), meta: `${formatNumber(summary.highConfidenceItems || 0)} registros com alta confiança.` },
      { tone: "type", label: "Tipo predominante", value: topType?.type || "Sem classificação", meta: `${formatNumber(topType?.count || 0)} registros no principal agrupamento.` },
      { tone: "institution", label: "Maior concentração institucional", value: topOrganization?.displayName || "Não identificado", meta: `${formatNumber(topOrganization?.count || 0)} registros no órgão com maior volume.` },
      { tone: "date", label: "Recência da base", value: formatDate(newestItem?.publishedAt), meta: "Data do registro público mais recente presente na base." },
    ];

    elements.executiveCards.innerHTML = cards.map((card) => `
      <article class="executive-card executive-card-${escapeHtml(card.tone)}">
        <span class="executive-card-label">${escapeHtml(card.label)}</span>
        <strong>${escapeHtml(card.value)}</strong>
        <p>${escapeHtml(card.meta)}</p>
      </article>
    `).join("");
  }

  function renderPriorities() {
    const items = getItems();
    const latestItems = sortByNewest(items).slice(0, 3);
    const highestValueItems = sortByHighestValue(items.filter((item) => Number(item.valueNumber || 0) > 0)).slice(0, 3);
    const itemsWithoutValue = sortByNewest(items.filter((item) => Number(item.valueNumber || 0) <= 0));
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
        items: itemsWithoutValue.slice(0, 3).map((item) => ({
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

  function renderInsights() {
    const items = getItems();
    const topYear = [...(state.payload?.yearSummary || [])].sort((a, b) => Number(b.count || 0) - Number(a.count || 0))[0];
    const latestWave = computeLatestWave(items);
    const topSupplier = buildSupplierRanking(items)[0];
    const missingValueCount = items.filter((item) => Number(item.valueNumber || 0) <= 0).length;
    const highConfidenceCount = items.filter((item) => String(item.confidence || "").toLowerCase() === "alta").length;
    const missingValueRate = items.length ? missingValueCount / items.length : 0;
    const highConfidenceRate = items.length ? highConfidenceCount / items.length : 0;

    elements.insightsMeta.textContent = "Leituras automáticas para ajudar quem abre o painel pela primeira vez.";

    const cards = [
      { tone: "year", label: "Ano com maior volume", value: topYear ? `${topYear.year}` : "Sem recorte", summary: topYear ? `${formatNumber(topYear.count)} registros concentrados no ano com maior incidência da base.` : "Sem série histórica suficiente para comparação.", actionType: "year", actionValue: topYear?.year || "", actionLabel: topYear ? `Filtrar ${topYear.year}` : "" },
      { tone: "wave", label: "Última edição relevante", value: latestWave ? `Edição ${latestWave.edition}` : "Sem edição", summary: latestWave ? `${formatNumber(latestWave.count)} registro(s) na publicação mais recente, em ${formatDate(latestWave.publishedAt)}.` : "Sem publicação recente identificada.", actionType: "search", actionValue: latestWave?.edition || "", actionLabel: latestWave ? "Buscar edição" : "" },
      { tone: "supplier", label: "Fornecedor mais recorrente", value: topSupplier ? truncateText(topSupplier.name, 30) : "Não identificado", summary: topSupplier ? `${formatNumber(topSupplier.count)} registro(s) associados e ${formatCurrency(topSupplier.totalValue || 0)} em valor identificado.` : "A base ainda não tem recorrência suficiente de fornecedores identificados.", actionType: "search", actionValue: topSupplier?.name || "", actionLabel: topSupplier ? "Buscar fornecedor" : "" },
      { tone: "missing", label: "Registros sem valor", value: formatPercent(missingValueRate), summary: `${formatNumber(missingValueCount)} registro(s) ainda sem valor consolidado para leitura rápida.`, actionType: "preset", actionValue: "missingValue", actionLabel: "Ver sem valor" },
      { tone: "confidence", label: "Leitura de alta confiança", value: formatPercent(highConfidenceRate), summary: `${formatNumber(highConfidenceCount)} registro(s) com leitura considerada mais consistente na base.`, actionType: "confidence", actionValue: "alta", actionLabel: "Filtrar alta confiança" },
    ];

    elements.insightCards.innerHTML = cards.map((card) => `
      <article class="insight-card insight-card-${escapeHtml(card.tone)}">
        <span class="insight-card-label">${escapeHtml(card.label)}</span>
        <strong>${escapeHtml(card.value)}</strong>
        <p>${escapeHtml(card.summary)}</p>
        ${card.actionLabel ? `<button class="insight-action" type="button" data-insight-action="${escapeHtml(card.actionType)}" data-insight-value="${escapeHtml(card.actionValue)}">${escapeHtml(card.actionLabel)}</button>` : ""}
      </article>
    `).join("");

    [...elements.insightCards.querySelectorAll("[data-insight-action]")].forEach((button) => {
      button.addEventListener("click", () => applyInsightAction(button.dataset.insightAction, button.dataset.insightValue));
    });
  }

  function renderQuickFilters() {
    const latestYear = getLatestYearEntry();
    const topType = getTopTypeEntry();
    const highestValueItem = sortByHighestValue(getItems().filter((item) => Number(item.valueNumber || 0) > 0))[0];
    const highConfidenceCount = getItems().filter((item) => String(item.confidence || "").toLowerCase() === "alta").length;
    const missingValueCount = getItems().filter((item) => Number(item.valueNumber || 0) <= 0).length;
    const presets = [
      { id: "recent", label: "Mais recentes", meta: "Ordem cronológica padrão" },
      { id: "latestYear", label: latestYear ? `Ano ${latestYear.year}` : "Ano mais recente", meta: latestYear ? `${formatNumber(latestYear.count)} registro(s)` : "Sem recorte anual" },
      { id: "topType", label: "Tipo predominante", meta: topType ? `${topType.type} (${formatNumber(topType.count)})` : "Sem classificação" },
      { id: "highestValue", label: "Maiores valores", meta: highestValueItem ? `Até ${highestValueItem.value || formatCurrency(highestValueItem.valueNumber)}` : "Sem valor identificado" },
      { id: "highConfidence", label: "Alta confiança", meta: `${formatNumber(highConfidenceCount)} registro(s)` },
      { id: "missingValue", label: "Sem valor", meta: `${formatNumber(missingValueCount)} registro(s)` },
    ];

    elements.quickFilters.innerHTML = presets.map((preset) => `
      <button class="quick-filter-button quick-filter-button-${escapeHtml(preset.id)} ${state.activePreset === preset.id ? "active" : ""}" data-preset="${escapeHtml(preset.id)}" type="button">
        <span>${escapeHtml(preset.label)}</span>
        <strong>${escapeHtml(preset.meta)}</strong>
      </button>
    `).join("");

    [...elements.quickFilters.querySelectorAll("[data-preset]")].forEach((button) => {
      button.addEventListener("click", () => applyQuickPreset(button.dataset.preset));
    });
  }

  function renderSelectionSummary() {
    const visibleRows = state.filteredItems.slice(0, state.visibleCount);
    const summary = computeSelectionSummary(state.filteredItems);
    const cards = [
      { label: "Recorte atual", value: formatNumber(state.filteredItems.length), meta: `${formatNumber(visibleRows.length)} registro(s) já visíveis na tela.` },
      { label: "Valor conhecido", value: formatCurrency(summary.totalValue), meta: `${formatNumber(summary.withValue)} registro(s) com valor informado no recorte.` },
      { label: "Com fornecedor", value: formatNumber(summary.withSupplier), meta: "Quantidade de registros com fornecedor identificado." },
      { label: "Alta confiança", value: state.filteredItems.length ? formatPercent(summary.highConfidence / state.filteredItems.length) : "0%", meta: `${formatNumber(summary.highConfidence)} registro(s) classificados com maior consistência.` },
    ];

    elements.selectionSummary.innerHTML = cards.map((card) => `
      <article class="selection-card">
        <span class="selection-card-label">${escapeHtml(card.label)}</span>
        <strong>${escapeHtml(card.value)}</strong>
        <p>${escapeHtml(card.meta)}</p>
      </article>
    `).join("");
  }

  function renderActiveFiltersSummary({ query, type, year, administration, confidence }) {
    const parts = [];
    const hasManualFilter = Boolean(query || type || year || administration || confidence || state.specialFilter === "missingValue" || state.sortMode !== "newest" || state.activePreset);
    if (!hasManualFilter) {
      elements.activeFilters.textContent = "Sem filtros ativos. A lista abaixo segue do mais novo para o mais antigo.";
      return;
    }

    if (state.activePreset) parts.push(`Atalho: ${getPresetLabel(state.activePreset)}`);
    if (query) parts.push(`Busca: ${query}`);
    if (type) parts.push(`Tipo: ${type}`);
    if (year) parts.push(`Ano: ${year}`);
    if (administration) parts.push(`Gestão: ${formatAdministrationLabel(administration)}`);
    if (confidence) parts.push(`Confiabilidade: ${formatConfidenceLabel(confidence)}`);
    if (state.specialFilter === "missingValue") parts.push("Somente registros sem valor informado");
    parts.push(`Ordenação: ${SORT_LABELS[state.sortMode] || SORT_LABELS.newest}`);
    elements.activeFilters.textContent = parts.join(" | ");
  }

  function renderTypePills() {
    const rows = state.payload?.typeSummary || [];
    elements.typePills.innerHTML = rows.map((row) => `
      <button class="type-pill ${state.activeType === row.type ? "active" : ""}" data-type="${escapeHtml(row.type)}" type="button">
        ${escapeHtml(row.type)} (${escapeHtml(row.count)})
      </button>
    `).join("");

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
    elements.resultsMeta.textContent = `${visibleRows.length} de ${state.filteredItems.length} registro(s) exibidos, ordenados por ${SORT_LABELS[state.sortMode] || SORT_LABELS.newest}.`;

    if (!visibleRows.length) {
      elements.contractsList.innerHTML = `<div class="empty-card">Nenhum contrato corresponde aos filtros atuais.</div>`;
      elements.loadMore.hidden = true;
      return;
    }

    elements.contractsList.innerHTML = visibleRows.map((item) => {
      const confidenceKey = String(item.confidence || "").toLowerCase();
      const confidenceLabel = formatConfidenceLabel(confidenceKey);
      return `
        <article class="contract-card">
          <div class="contract-card-main">
            <div class="contract-head">
              <div>
                <div class="badge-row">
                  <span class="badge">${escapeHtml(item.type || "Ato")}</span>
                  <span class="badge neutral">${escapeHtml(item.recordClass || "Registro")}</span>
                  ${confidenceKey ? `<span class="badge trust trust-${escapeHtml(confidenceKey || "media")}">${escapeHtml(confidenceLabel)}</span>` : ""}
                </div>
                <h3>${escapeHtml(item.title || "Ato contratual")}</h3>
              </div>
              <div class="meta-row">
                <span class="meta-chip">${escapeHtml(`Edição ${item.edition || "-"}`)}</span>
                <span class="meta-chip">${escapeHtml(formatDate(item.publishedAt))}</span>
              </div>
            </div>
            <p class="contract-summary">${escapeHtml(item.summary || "Sem resumo consolidado.")}</p>
          </div>
          <aside class="contract-side">
            <div class="contract-side-item">
              <span>Órgão</span>
              <strong>${escapeHtml(item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere))}</strong>
            </div>
            <div class="contract-side-item">
              <span>Fornecedor</span>
              <strong>${escapeHtml(item.contractor || "Não identificado")}</strong>
            </div>
            <div class="contract-side-item">
              <span>Valor</span>
              <strong>${escapeHtml(item.value || "Não informado")}</strong>
            </div>
          </aside>
          <div class="action-row">
            ${item.viewUrl ? `<a class="action-link" href="${escapeHtml(item.viewUrl)}" target="_blank" rel="noopener noreferrer">Abrir edição oficial</a>` : ""}
          </div>
        </article>
      `;
    }).join("");

    elements.loadMore.hidden = visibleRows.length >= state.filteredItems.length;
  }

  function resetAllFilters() {
    if (elements.heroSearchInput) elements.heroSearchInput.value = "";
    elements.searchInput.value = "";
    elements.typeSelect.value = "";
    elements.yearSelect.value = "";
    elements.administrationSelect.value = "";
    elements.confidenceSelect.value = "";
    elements.sortSelect.value = "newest";
    state.activeType = "";
    state.sortMode = "newest";
    state.specialFilter = "";
    state.activePreset = "";
  }

  function updateFilters({ resetVisible = false } = {}) {
    const rawQuery = elements.searchInput.value.trim();
    const query = normalizeForSearch(rawQuery);
    const type = elements.typeSelect.value;
    const year = elements.yearSelect.value;
    const administration = elements.administrationSelect.value;
    const confidence = elements.confidenceSelect.value;
    state.activeType = type;
    state.sortMode = elements.sortSelect.value || "newest";

    const filtered = getItems().filter((item) => {
      const organizationLabel = item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere);
      const haystack = normalizeForSearch([item.title, item.organization, organizationLabel, item.contractor, item.summary, item.type, item.edition].join(" "));
      if (query && !haystack.includes(query)) return false;
      if (type && item.type !== type) return false;
      if (year && !(item.publishedAt || "").startsWith(year)) return false;
      if (administration && getAdministrationValueForDate(item.publishedAt) !== administration) return false;
      if (confidence && String(item.confidence || "").toLowerCase() !== confidence) return false;
      if (state.specialFilter === "missingValue" && Number(item.valueNumber || 0) > 0) return false;
      return true;
    });

    state.filteredItems = sortItems(filtered);
    if (resetVisible) state.visibleCount = PAGE_SIZE;
    syncUrlState();
    renderQuickFilters();
    renderActiveFiltersSummary({ query: rawQuery, type, year, administration, confidence });
    renderSelectionSummary();
    renderTypePills();
    renderContracts();
  }

  function applyQuickPreset(presetId) {
    const latestYear = getLatestYearEntry();
    const topType = getTopTypeEntry();
    resetAllFilters();
    state.activePreset = presetId;

    if (presetId === "latestYear" && latestYear?.year) {
      elements.yearSelect.value = String(latestYear.year);
    }
    else if (presetId === "topType" && topType?.type) {
      elements.typeSelect.value = String(topType.type);
    }
    else if (presetId === "highestValue") {
      elements.sortSelect.value = "highestValue";
    }
    else if (presetId === "highConfidence") {
      elements.confidenceSelect.value = "alta";
    }
    else if (presetId === "missingValue") {
      state.specialFilter = "missingValue";
    }

    updateFilters({ resetVisible: true });
  }

  function applyInsightAction(action, value) {
    if (!action || !value) return;
    resetAllFilters();
    applyView("consulta", { sync: false });
    if (action === "year") {
      elements.yearSelect.value = value;
      const administrationValue = getAdministrationValueForDate(`${value}-01-01`);
      if (administrationValue) elements.administrationSelect.value = administrationValue;
    }
    else if (action === "search") {
      elements.searchInput.value = value;
      if (elements.heroSearchInput) elements.heroSearchInput.value = value;
    }
    else if (action === "confidence") {
      elements.confidenceSelect.value = value;
    }
    else if (action === "preset") {
      state.specialFilter = value === "missingValue" ? "missingValue" : "";
      state.activePreset = value;
    }
    updateFilters({ resetVisible: true });
    document.getElementById("section-controls")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function handleCopyLink() {
    syncUrlState();
    const copied = await copyText(window.location.href);
    setFeedback(copied ? "Link da consulta copiado." : "Não foi possível copiar o link automaticamente.", copied ? "success" : "warning");
  }

  function handleExportCsv() {
    if (!state.filteredItems.length) {
      setFeedback("Não há registros no recorte atual para exportar.", "warning");
      return;
    }

    const headers = ["Data", "Edição", "Tipo", "Classe", "Órgão", "Fornecedor", "Valor", "Confiabilidade", "Título", "Resumo", "Link oficial"];
    const rows = state.filteredItems.map((item) => [
      formatDate(item.publishedAt),
      item.edition || "-",
      item.type || "",
      item.recordClass || "",
      item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere),
      item.contractor || "",
      item.value || "",
      formatConfidenceLabel(item.confidence),
      item.title || "",
      item.summary || "",
      item.viewUrl || "",
    ]);

    const csv = ["\uFEFF" + headers.map(toCsvValue).join(";"), ...rows.map((row) => row.map(toCsvValue).join(";"))].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `contratos-iguape-${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    setFeedback("Arquivo CSV exportado com o recorte atual.", "success");
  }

  async function loadDashboard() {
    const response = await fetch("./data/dashboard.json", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Falha ao carregar a base pública: ${response.status}`);
    }

    state.payload = normalizePayload(await response.json());
    populateSelect(elements.typeSelect, (state.payload.typeSummary || []).map((item) => ({ value: item.type, label: item.type })), "Todos os tipos");
    populateSelect(elements.yearSelect, (state.payload.yearSummary || []).map((item) => ({ value: item.year, label: `${item.year} (${item.count})` })), "Todos os anos");
    populateSelect(elements.administrationSelect, buildAdministrationEntries(getItems()), "Todas as gestões");
    populateSelect(elements.confidenceSelect, buildConfidenceEntries(getItems()), "Todas as leituras");

    applyInitialUrlState();
    applyView(state.currentView, { sync: false });
    renderMetrics();
    renderExecutiveSummary();
    renderPriorities();
    renderInsights();
    renderQuickFilters();
    updateFilters({ resetVisible: true });
  }

  function wireEvents() {
    if (elements.heroSearchForm && elements.heroSearchInput) {
      elements.heroSearchForm.addEventListener("submit", (event) => {
        event.preventDefault();
        elements.searchInput.value = elements.heroSearchInput.value.trim();
        state.activePreset = "";
        applyView("consulta", { sync: false });
        updateFilters({ resetVisible: true });
        document.getElementById("section-controls")?.scrollIntoView({ behavior: "smooth", block: "start" });
        elements.searchInput.focus({ preventScroll: true });
      });
    }

    elements.searchInput.addEventListener("input", () => {
      if (elements.heroSearchInput) elements.heroSearchInput.value = elements.searchInput.value;
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.typeSelect.addEventListener("change", () => {
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.yearSelect.addEventListener("change", () => {
      if (elements.yearSelect.value) {
        const administrationValue = getAdministrationValueForDate(`${elements.yearSelect.value}-01-01`);
        if (administrationValue) elements.administrationSelect.value = administrationValue;
      }
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.administrationSelect.addEventListener("change", () => {
      if (!administrationIncludesYear(elements.administrationSelect.value, elements.yearSelect.value)) {
        elements.yearSelect.value = "";
      }
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.confidenceSelect.addEventListener("change", () => {
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.sortSelect.addEventListener("change", () => {
      state.activePreset = "";
      updateFilters({ resetVisible: true });
    });
    elements.resetFilters.addEventListener("click", () => {
      resetAllFilters();
      updateFilters({ resetVisible: true });
    });
    elements.copyLink.addEventListener("click", () => {
      handleCopyLink();
    });
    elements.exportCsv.addEventListener("click", () => {
      handleExportCsv();
    });
    elements.loadMore.addEventListener("click", () => {
      state.visibleCount += PAGE_SIZE;
      renderSelectionSummary();
      renderContracts();
    });
    elements.viewButtons.forEach((button) => {
      button.addEventListener("click", () => {
        applyView(button.dataset.viewButton, { scroll: true });
      });
    });
    elements.viewLinks.forEach((link) => {
      link.addEventListener("click", (event) => {
        const targetView = link.dataset.viewTarget;
        if (!VALID_VIEWS.has(targetView)) return;
        const href = link.getAttribute("href") || "";
        if (href.startsWith("#")) event.preventDefault();
        applyView(targetView, { sync: false });
        if (href.startsWith("#")) {
          document.querySelector(href)?.scrollIntoView({ behavior: "smooth", block: "start" });
          syncUrlState();
        }
      });
    });
  }

  wireEvents();
  loadDashboard().catch((error) => {
    const message = escapeHtml(error.message);
    elements.headlineMetrics.innerHTML = `<div class="empty-card">${message}</div>`;
    elements.executiveOverview.innerHTML = `<div class="empty-card">${message}</div>`;
    elements.executiveCards.innerHTML = "";
    elements.executiveMeta.textContent = "Não foi possível consolidar o panorama gerencial.";
    elements.priorityCards.innerHTML = "";
    elements.prioritiesMeta.textContent = "Não foi possível consolidar os pontos de atenção.";
    elements.insightCards.innerHTML = "";
    elements.insightsMeta.textContent = "Não foi possível consolidar os insights automáticos.";
    elements.quickFilters.innerHTML = "";
    elements.selectionSummary.innerHTML = `<div class="empty-card">${message}</div>`;
    elements.activeFilters.textContent = "Não foi possível consolidar os atalhos de navegação.";
    elements.contractsList.innerHTML = `<div class="empty-card">${message}</div>`;
    elements.resultsMeta.textContent = "Não foi possível carregar a base pública.";
  });
})();
