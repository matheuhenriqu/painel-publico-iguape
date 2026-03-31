(() => {
  const PAGE_SIZE = 30;
  const state = {
    payload: null,
    filteredItems: [],
    visibleCount: PAGE_SIZE,
    activeType: "",
  };

  const elements = {
    headlineMetrics: document.getElementById("headline-metrics"),
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

  function formatCurrency(value) {
    return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(Number(value || 0));
  }

  function formatDate(value, includeTime = false) {
    if (!value) return "Sem data";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return new Intl.DateTimeFormat("pt-BR", includeTime ? { dateStyle: "medium", timeStyle: "short" } : { dateStyle: "medium" }).format(date);
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
    if (!normalizedName) return "Nao identificado";

    const governmentLabel = getGovernmentLabel(sphere);
    return governmentLabel ? `${normalizedName} - ${governmentLabel}` : normalizedName;
  }

  function getItems() {
    return state.payload?.items || [];
  }

  function sortByNewest(items) {
    return [...items].sort((a, b) => new Date(b.publishedAt || 0) - new Date(a.publishedAt || 0) || Number(b.edition || 0) - Number(a.edition || 0));
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
      { label: "Registros", value: summary.totalItems || 0, meta: "atos organizados" },
      { label: "Valor identificado", value: formatCurrency(summary.totalValue || 0), meta: "somatorio extraido" },
      { label: "Alta confianca", value: summary.highConfidenceItems || 0, meta: "itens consistentes" },
    ];
    elements.headlineMetrics.innerHTML = cards.map((card) => `<article class="metric-card"><strong>${escapeHtml(card.value)}</strong><span>${escapeHtml(card.label)}</span><span>${escapeHtml(card.meta)}</span></article>`).join("");
    elements.footerCopy.textContent = `Atualizado em ${formatDate(state.payload?.generatedAt, true)}. ${summary.analyzedDiaryCount || 0} edicoes analisadas e ${summary.uniqueSuppliers || 0} fornecedores identificados.`;
  }

  function renderTypePills() {
    const rows = state.payload?.typeSummary || [];
    elements.typePills.innerHTML = rows.map((row) => `<button class="type-pill ${state.activeType === row.type ? "active" : ""}" data-type="${escapeHtml(row.type)}" type="button">${escapeHtml(row.type)} (${escapeHtml(row.count)})</button>`).join("");
    [...elements.typePills.querySelectorAll("[data-type]")].forEach((button) => {
      button.addEventListener("click", () => {
        state.activeType = state.activeType === button.dataset.type ? "" : button.dataset.type;
        elements.typeSelect.value = state.activeType;
        updateFilters({ resetVisible: true });
      });
    });
  }

  function renderContracts() {
    const visibleRows = state.filteredItems.slice(0, state.visibleCount);
    elements.resultsMeta.textContent = `${visibleRows.length} de ${state.filteredItems.length} registro(s) exibidos, sempre do mais novo para o mais antigo.`;
    if (!visibleRows.length) {
      elements.contractsList.innerHTML = `<div class="empty-card">Nenhum contrato corresponde aos filtros atuais.</div>`;
      elements.loadMore.hidden = true;
      return;
    }

    elements.contractsList.innerHTML = visibleRows.map((item) => `<article class="contract-card"><div class="contract-head"><div><div class="badge-row"><span class="badge">${escapeHtml(item.type || "Ato")}</span><span class="badge neutral">${escapeHtml(item.recordClass || "Registro")}</span></div><h3>${escapeHtml(item.title || "Ato contratual")}</h3></div><div class="meta-row"><span class="meta-chip">${escapeHtml(`Edicao ${item.edition || "-"}`)}</span><span class="meta-chip">${escapeHtml(formatDate(item.publishedAt))}</span></div></div><div class="meta-line"><strong>Orgao:</strong> ${escapeHtml(item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere))}</div><div class="meta-line"><strong>Fornecedor:</strong> ${escapeHtml(item.contractor || "Nao identificado")}</div><div class="meta-line"><strong>Valor:</strong> ${escapeHtml(item.value || "Nao informado")}</div><p class="contract-summary">${escapeHtml(item.summary || "Sem resumo consolidado.")}</p><div class="action-row">${item.viewUrl ? `<a class="action-link" href="${escapeHtml(item.viewUrl)}" target="_blank" rel="noopener noreferrer">Abrir edicao oficial</a>` : ""}</div></article>`).join("");
    elements.loadMore.hidden = visibleRows.length >= state.filteredItems.length;
  }

  function updateFilters({ resetVisible = false } = {}) {
    const query = elements.searchInput.value.trim().toLowerCase();
    const type = elements.typeSelect.value;
    const year = elements.yearSelect.value;
    state.activeType = type;
    state.filteredItems = sortByNewest(getItems().filter((item) => {
      const organizationLabel = item.organizationDisplay || formatOrganizationLabel(item.organization, item.organizationSphere);
      const haystack = [item.title, item.organization, organizationLabel, item.contractor, item.summary, item.type, item.edition].join(" ").toLowerCase();
      if (query && !haystack.includes(query)) return false;
      if (type && item.type !== type) return false;
      if (year && !(item.publishedAt || "").startsWith(year)) return false;
      return true;
    }));
    if (resetVisible) state.visibleCount = PAGE_SIZE;
    renderTypePills();
    renderContracts();
  }

  async function loadDashboard() {
    const response = await fetch("./data/dashboard.json", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Falha ao carregar a base publica: ${response.status}`);
    }
    state.payload = await response.json();
    const typeOptions = (state.payload.typeSummary || []).map((item) => ({ value: item.type, label: item.type }));
    const yearOptions = (state.payload.yearSummary || []).map((item) => ({ value: item.year, label: `${item.year} (${item.count})` }));
    populateSelect(elements.typeSelect, typeOptions, "Todos os tipos");
    populateSelect(elements.yearSelect, yearOptions, "Todos os anos");
    renderMetrics();
    updateFilters({ resetVisible: true });
  }

  function wireEvents() {
    elements.searchInput.addEventListener("input", () => updateFilters({ resetVisible: true }));
    elements.typeSelect.addEventListener("change", () => updateFilters({ resetVisible: true }));
    elements.yearSelect.addEventListener("change", () => updateFilters({ resetVisible: true }));
    elements.resetFilters.addEventListener("click", () => {
      elements.searchInput.value = "";
      elements.typeSelect.value = "";
      elements.yearSelect.value = "";
      state.activeType = "";
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
    elements.contractsList.innerHTML = `<div class="empty-card">${escapeHtml(error.message)}</div>`;
    elements.resultsMeta.textContent = "Nao foi possivel carregar a base publica.";
  });
})();
