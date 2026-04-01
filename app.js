(() => {
  const PAGE_SIZE = 18;
  const DEFAULT_VIEW = "overview";
  const VALID_VIEWS = new Set(["overview", "alerts", "contracts"]);
  const DEFAULT_FILTERS = {
    query: "",
    organization: "",
    administration: "",
    vigency: "todos",
    management: "todos",
    source: "todos",
    scope: "atuais",
  };

  const LABELS = {
    vigency: {
      todos: "Todas as situações",
      vigente_confirmado: "Vigência confirmada",
      vigente_inferido: "Vigência inferida",
      em_acompanhamento: "Em acompanhamento",
      encerrado: "Encerrado",
      sem_sinal_atual: "Sem sinal atual",
    },
    management: {
      todos: "Todas as situações",
      completos: "Responsáveis completos",
      sem_gestor: "Sem gestor atual",
      sem_fiscal: "Sem fiscal atual",
      sem_gestor_e_fiscal: "Sem gestor e fiscal",
      revisao: "Precisa revisão",
      exoneracao: "Sinal de exoneração",
    },
    source: {
      todos: "Todas as origens",
      cruzado: "Diário + portal",
      somente_diario: "Somente Diário",
      somente_portal: "Somente portal",
    },
    scope: {
      atuais: "Somente contratos atuais",
      todos: "Todos os contratos monitorados",
    },
  };

  const PRESETS = {
    semGestorEFiscal: {
      label: "Sem gestor e fiscal",
      description: "Contratos atuais sem nenhum responsável confirmado.",
      filters: { management: "sem_gestor_e_fiscal", scope: "atuais" },
    },
    semGestor: {
      label: "Sem gestor",
      description: "Contratos atuais em que falta o gestor.",
      filters: { management: "sem_gestor", scope: "atuais" },
    },
    semFiscal: {
      label: "Sem fiscal",
      description: "Contratos atuais em que falta o fiscal.",
      filters: { management: "sem_fiscal", scope: "atuais" },
    },
    somenteDiario: {
      label: "Somente Diário",
      description: "Contratos atuais ainda sem correspondência automática no portal.",
      filters: { source: "somente_diario", scope: "atuais" },
    },
    completos: {
      label: "Responsáveis completos",
      description: "Contratos atuais com gestor e fiscal identificados.",
      filters: { management: "completos", scope: "atuais" },
    },
  };

  const state = {
    payload: null,
    view: DEFAULT_VIEW,
    filters: { ...DEFAULT_FILTERS },
    visibleCount: PAGE_SIZE,
  };

  const elements = {
    updatedAt: document.getElementById("updated-at"),
    heroSummary: document.getElementById("hero-summary"),
    heroCallout: document.getElementById("hero-callout"),
    summaryCards: document.getElementById("summary-cards"),
    methodSummary: document.getElementById("method-summary"),
    methodNotes: document.getElementById("method-notes"),
    statusGrid: document.getElementById("status-grid"),
    priorityGroups: document.getElementById("priority-groups"),
    organizationSummary: document.getElementById("organization-summary"),
    alertRecords: document.getElementById("alert-records"),
    searchInput: document.getElementById("search-input"),
    scopeSelect: document.getElementById("scope-select"),
    organizationSelect: document.getElementById("organization-select"),
    administrationSelect: document.getElementById("administration-select"),
    vigencySelect: document.getElementById("vigency-select"),
    managementSelect: document.getElementById("management-select"),
    sourceSelect: document.getElementById("source-select"),
    quickPresets: document.getElementById("quick-presets"),
    resultsMeta: document.getElementById("results-meta"),
    activeFilterSummary: document.getElementById("active-filter-summary"),
    recordList: document.getElementById("record-list"),
    loadMore: document.getElementById("load-more"),
    clearFilters: document.getElementById("clear-filters"),
    footerCopy: document.getElementById("footer-copy"),
    viewButtons: [...document.querySelectorAll("[data-view-button]")],
    viewPanels: [...document.querySelectorAll("[data-view-panel]")],
  };

  function readUrlState() {
    const params = new URLSearchParams(window.location.search);
    const view = VALID_VIEWS.has(params.get("view")) ? params.get("view") : DEFAULT_VIEW;
    return {
      view,
      filters: {
        query: params.get("q") || "",
        organization: params.get("org") || "",
        administration: params.get("adm") || "",
        vigency: params.get("vig") || DEFAULT_FILTERS.vigency,
        management: params.get("mgmt") || DEFAULT_FILTERS.management,
        source: params.get("src") || DEFAULT_FILTERS.source,
        scope: params.get("scope") || DEFAULT_FILTERS.scope,
      },
    };
  }

  function syncUrlState() {
    const params = new URLSearchParams();
    if (state.view !== DEFAULT_VIEW) params.set("view", state.view);
    if (state.filters.query) params.set("q", state.filters.query);
    if (state.filters.organization) params.set("org", state.filters.organization);
    if (state.filters.administration) params.set("adm", state.filters.administration);
    if (state.filters.vigency !== DEFAULT_FILTERS.vigency) params.set("vig", state.filters.vigency);
    if (state.filters.management !== DEFAULT_FILTERS.management) params.set("mgmt", state.filters.management);
    if (state.filters.source !== DEFAULT_FILTERS.source) params.set("src", state.filters.source);
    if (state.filters.scope !== DEFAULT_FILTERS.scope) params.set("scope", state.filters.scope);
    const query = params.toString();
    window.history.replaceState({}, "", query ? `${window.location.pathname}?${query}` : window.location.pathname);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function normalizeText(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function formatDate(value) {
    if (!value) return "Não informado";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return new Intl.DateTimeFormat("pt-BR", { dateStyle: "medium" }).format(date);
  }

  function formatDateTime(value) {
    if (!value) return "Não informado";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return new Intl.DateTimeFormat("pt-BR", { dateStyle: "medium", timeStyle: "short" }).format(date);
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("pt-BR").format(Number(value || 0));
  }

  function formatCurrency(value) {
    if (!value || Number(value) <= 0) return "Sem valor consolidado";
    return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(Number(value));
  }

  function truncateText(value, limit = 260) {
    const text = String(value || "").replace(/\s+/g, " ").trim();
    if (text.length <= limit) return text;
    return `${text.slice(0, limit - 3).trim()}...`;
  }

  function getRecords() {
    return state.payload?.records || [];
  }

  function getCurrentRecords() {
    return getRecords().filter((record) => record.vigency?.isCurrent);
  }

  function getLabel(group, value) {
    return LABELS[group]?.[value] || value || "Não informado";
  }

  function getToneClass(record) {
    if ((record.alertWeight || 0) >= 3) return "record-card--critical";
    if ((record.alertWeight || 0) >= 2) return "record-card--warning";
    return "";
  }

  function getBadgeToneByManagement(value) {
    if (value === "completos") return "success";
    if (value === "sem_gestor_e_fiscal" || value === "exoneracao") return "danger";
    if (value === "sem_gestor" || value === "sem_fiscal" || value === "revisao") return "warning";
    return "primary";
  }

  function getBadgeToneByVigency(value) {
    if (value === "vigente_confirmado") return "success";
    if (value === "vigente_inferido" || value === "em_acompanhamento") return "warning";
    if (value === "encerrado") return "danger";
    return "primary";
  }

  function getBadgeToneBySource(value) {
    if (value === "cruzado") return "success";
    if (value === "somente_diario" || value === "somente_portal") return "warning";
    return "primary";
  }

  function getUniqueAlerts(alerts) {
    const seen = new Set();
    return (alerts || []).filter((alert) => {
      const key = `${alert.title}|${alert.severity}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function matchesQuery(record, query) {
    if (!query) return true;
    const haystack = [
      record.contractNumber,
      record.organization,
      record.supplier,
      record.object,
      record.managementSummary,
      record.manager?.name,
      record.manager?.role,
      record.inspector?.name,
      record.inspector?.role,
      ...(record.alerts || []).map((alert) => alert.title),
    ].join(" ");
    return normalizeText(haystack).includes(normalizeText(query));
  }

  function getFilteredRecords() {
    return getRecords().filter((record) => {
      if (state.filters.scope === "atuais" && !record.vigency?.isCurrent) return false;
      if (state.filters.organization && record.organization !== state.filters.organization) return false;
      if (state.filters.administration && record.administration !== state.filters.administration) return false;
      if (state.filters.vigency !== "todos" && record.vigency?.state !== state.filters.vigency) return false;
      if (state.filters.management !== "todos" && record.managementState !== state.filters.management) return false;
      if (state.filters.source !== "todos" && record.sourceStatus !== state.filters.source) return false;
      if (!matchesQuery(record, state.filters.query)) return false;
      return true;
    });
  }

  function setView(view) {
    state.view = VALID_VIEWS.has(view) ? view : DEFAULT_VIEW;
    elements.viewButtons.forEach((button) => {
      button.classList.toggle("active", button.dataset.viewButton === state.view);
    });
    elements.viewPanels.forEach((panel) => {
      panel.classList.toggle("hidden", panel.dataset.viewPanel !== state.view);
    });
    syncUrlState();
  }

  function syncControls() {
    elements.searchInput.value = state.filters.query;
    elements.scopeSelect.value = state.filters.scope;
    elements.organizationSelect.value = state.filters.organization;
    elements.administrationSelect.value = state.filters.administration;
    elements.vigencySelect.value = state.filters.vigency;
    elements.managementSelect.value = state.filters.management;
    elements.sourceSelect.value = state.filters.source;
  }

  function populateSelect(select, values, groupName, emptyValue = "todos") {
    const previous = select.value;
    select.innerHTML = "";
    const allOption = document.createElement("option");
    allOption.value = emptyValue;
    allOption.textContent = getLabel(groupName, emptyValue);
    select.appendChild(allOption);

    values
      .filter((value) => value !== emptyValue)
      .forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = getLabel(groupName, value);
      select.appendChild(option);
      });

    if ([...select.options].some((option) => option.value === previous)) {
      select.value = previous;
    }
  }

  function renderSummaryCards() {
    const summary = state.payload.summary;
    elements.summaryCards.innerHTML = [
      {
        label: "Contratos atuais",
        value: summary.contratosAtuais,
        detail: `${formatNumber(summary.vigentesInferidos)} com vigência inferida e ${formatNumber(summary.emAcompanhamento)} em acompanhamento no Diário.`,
      },
      {
        label: "Sem gestor e fiscal",
        value: summary.semGestorEFiscal,
        detail: "Casos em que nenhum responsável atual foi confirmado.",
      },
      {
        label: "Responsáveis completos",
        value: summary.comResponsaveisCompletos,
        detail: "Contratos atuais com gestor e fiscal identificados.",
      },
      {
        label: "Alertas críticos",
        value: summary.alertasCriticos,
        detail: `${formatNumber(summary.somenteDiario)} contratos atuais aparecem apenas no Diário Oficial.`,
      },
    ]
      .map(
        (card) => `
          <article class="metric-card">
            <span class="eyebrow">${escapeHtml(card.label)}</span>
            <strong class="metric-value">${formatNumber(card.value)}</strong>
            <span class="metric-detail">${escapeHtml(card.detail)}</span>
          </article>
        `
      )
      .join("");
  }

  function renderMethodology() {
    elements.updatedAt.textContent = `Atualizado em ${formatDateTime(state.payload.generatedAt)}`;
    elements.methodSummary.textContent = state.payload.methodology.summary;
    elements.methodNotes.innerHTML = state.payload.methodology.notes
      .map((note) => `<article class="note-card">${escapeHtml(note)}</article>`)
      .join("");
  }

  function renderHero() {
    const summary = state.payload.summary;
    elements.heroSummary.textContent =
      `A base atual acompanha ${formatNumber(summary.contratosAtuais)} contratos como atuais. ` +
      `${formatNumber(summary.semGestorEFiscal)} deles seguem sem gestor e fiscal confirmados, e ${formatNumber(summary.comResponsaveisCompletos)} já mostram os dois responsáveis identificados.`;

    elements.heroCallout.textContent =
      summary.vigentesConfirmados === 0
        ? "A maior parte da leitura atual depende de movimentação recente no Diário Oficial, porque o portal não confirma vigência ativa para quase toda a base cruzada."
        : `Há ${formatNumber(summary.vigentesConfirmados)} contrato(s) com vigência confirmada diretamente no portal oficial.`;
  }

  function renderStatusGrid() {
    const summary = state.payload.summary;
    const cards = [
      {
        label: "Vigência inferida",
        value: summary.vigentesInferidos,
        detail: "Contratos em que o prazo atual foi encontrado em documento ou termo.",
        className: "status-card--warning",
      },
      {
        label: "Em acompanhamento",
        value: summary.emAcompanhamento,
        detail: "Contratos tratados como atuais porque têm movimentação recente no Diário Oficial.",
        className: "status-card--warning",
      },
      {
        label: "Sem gestor atual",
        value: summary.semGestor,
        detail: "Contratos atuais em que o gestor não foi confirmado.",
        className: "status-card--danger",
      },
      {
        label: "Sem fiscal atual",
        value: summary.semFiscal,
        detail: "Contratos atuais em que o fiscal não foi confirmado.",
        className: "status-card--danger",
      },
    ];

    elements.statusGrid.innerHTML = cards
      .map(
        (card) => `
          <article class="status-card ${card.className}">
            <span class="eyebrow">${escapeHtml(card.label)}</span>
            <strong class="status-value">${formatNumber(card.value)}</strong>
            <span class="status-detail">${escapeHtml(card.detail)}</span>
          </article>
        `
      )
      .join("");
  }

  function sampleContracts(records) {
    return records
      .slice(0, 3)
      .map((record) => record.contractNumber)
      .join(" · ");
  }

  function renderPriorityGroups() {
    const currentRecords = getCurrentRecords();
    const groups = [
      {
        preset: "semGestorEFiscal",
        title: PRESETS.semGestorEFiscal.label,
        count: currentRecords.filter((record) => record.managementState === "sem_gestor_e_fiscal").length,
        sample: sampleContracts(currentRecords.filter((record) => record.managementState === "sem_gestor_e_fiscal")),
      },
      {
        preset: "semGestor",
        title: PRESETS.semGestor.label,
        count: currentRecords.filter((record) => record.managementState === "sem_gestor").length,
        sample: sampleContracts(currentRecords.filter((record) => record.managementState === "sem_gestor")),
      },
      {
        preset: "semFiscal",
        title: PRESETS.semFiscal.label,
        count: currentRecords.filter((record) => record.managementState === "sem_fiscal").length,
        sample: sampleContracts(currentRecords.filter((record) => record.managementState === "sem_fiscal")),
      },
      {
        preset: "somenteDiario",
        title: PRESETS.somenteDiario.label,
        count: currentRecords.filter((record) => record.sourceStatus === "somente_diario").length,
        sample: sampleContracts(currentRecords.filter((record) => record.sourceStatus === "somente_diario")),
      },
    ];

    elements.priorityGroups.innerHTML = groups
      .map(
        (group) => `
          <article class="priority-card">
            <span class="eyebrow">Atalho de consulta</span>
            <strong>${escapeHtml(group.title)}</strong>
            <span class="priority-count">${formatNumber(group.count)}</span>
            <p>${escapeHtml(PRESETS[group.preset].description)}</p>
            <div class="priority-sample">${escapeHtml(group.sample || "Sem exemplos prioritários nesta faixa.")}</div>
            <button type="button" data-preset="${escapeHtml(group.preset)}">Abrir lista</button>
          </article>
        `
      )
      .join("");
  }

  function renderOrganizationSummary() {
    const list = state.payload.organizationSummary || [];
    if (!list.length) {
      elements.organizationSummary.innerHTML = `<div class="empty-state">Sem órgãos para exibir.</div>`;
      return;
    }

    const maxCount = Math.max(...list.map((item) => item.count), 1);
    elements.organizationSummary.innerHTML = list
      .map(
        (item) => `
          <div class="organization-row">
            <div class="organization-head">
              <strong>${escapeHtml(item.organization)}</strong>
              <span>${formatNumber(item.count)} contrato(s)</span>
            </div>
            <div class="organization-bar">
              <span style="width: ${Math.max(10, (item.count / maxCount) * 100)}%"></span>
            </div>
          </div>
        `
      )
      .join("");
  }

  function renderQuickPresets() {
    elements.quickPresets.innerHTML = Object.entries(PRESETS)
      .map(([key, preset]) => {
        const active =
          Object.entries(preset.filters).every(([filterKey, filterValue]) => state.filters[filterKey] === filterValue) &&
          state.view === "contracts";
        return `
          <button class="filter-chip ${active ? "active" : ""}" type="button" data-preset="${escapeHtml(key)}">
            ${escapeHtml(preset.label)}
          </button>
        `;
      })
      .join("");
  }

  function getPersonDisplay(person) {
    if (person?.name) {
      return {
        title: person.name,
        subtitle: person.role || "Responsável identificado na base pública.",
      };
    }

    if (person?.needsReview) {
      return {
        title: "Precisa revisão",
        subtitle: "O nome extraído do Diário Oficial não foi tratado como confirmação válida.",
      };
    }

    return {
      title: "Não identificado",
      subtitle: "Sem responsável atual confirmado na leitura pública.",
    };
  }

  function renderBadges(record) {
    return [
      {
        tone: getBadgeToneByVigency(record.vigency?.state),
        label: getLabel("vigency", record.vigency?.state),
      },
      {
        tone: getBadgeToneByManagement(record.managementState),
        label: getLabel("management", record.managementState),
      },
      {
        tone: getBadgeToneBySource(record.sourceStatus),
        label: getLabel("source", record.sourceStatus),
      },
    ]
      .map(
        (badge) => `
          <span class="badge badge--${escapeHtml(badge.tone)}">${escapeHtml(badge.label)}</span>
        `
      )
      .join("");
  }

  function renderAlertPills(record) {
    const alerts = getUniqueAlerts(record.alerts).slice(0, 4);
    if (!alerts.length) return "";
    return `
      <ul class="alert-list">
        ${alerts
          .map(
            (alert) => `
              <li class="alert-pill alert-pill--${escapeHtml(alert.severity || "info")}">${escapeHtml(alert.title)}</li>
            `
          )
          .join("")}
      </ul>
    `;
  }

  function renderRecordCard(record) {
    const manager = getPersonDisplay(record.manager);
    const inspector = getPersonDisplay(record.inspector);
    const diaryLink = record.links?.diary
      ? `<a href="${escapeHtml(record.links.diary)}" target="_blank" rel="noopener noreferrer">Abrir Diário Oficial</a>`
      : "";
    const portalLink = record.links?.portal
      ? `<a href="${escapeHtml(record.links.portal)}" target="_blank" rel="noopener noreferrer">Abrir portal</a>`
      : "";

    return `
      <article class="record-card ${getToneClass(record)}">
        <div class="record-head">
          <div class="record-heading">
            <span class="record-number">${escapeHtml(record.contractNumber || "Contrato sem número")}</span>
            <h3>${escapeHtml(record.organization || "Órgão não identificado")}</h3>
            <span class="record-summary">${escapeHtml(record.managementSummary || "Sem resumo de gestão.")}</span>
          </div>
          <div class="badge-row">
            ${renderBadges(record)}
          </div>
        </div>

        <p class="record-object">${escapeHtml(truncateText(record.object || "Sem objeto resumido disponível."))}</p>

        <div class="record-meta">
          <div class="meta-block">
            <span>Gestor</span>
            <strong>${escapeHtml(manager.title)}</strong>
            <small>${escapeHtml(manager.subtitle)}</small>
          </div>
          <div class="meta-block">
            <span>Fiscal</span>
            <strong>${escapeHtml(inspector.title)}</strong>
            <small>${escapeHtml(inspector.subtitle)}</small>
          </div>
          <div class="meta-block">
            <span>Vigência</span>
            <strong>${escapeHtml(record.vigency?.label || "Sem vigência identificada")}</strong>
            <small>${escapeHtml(record.vigency?.sourceLabel || "Sem fonte de vigência")}</small>
          </div>
          <div class="meta-block">
            <span>Última movimentação</span>
            <strong>${escapeHtml(formatDate(record.managementActAt || record.publishedAt))}</strong>
            <small>${escapeHtml(record.lastMovementTitle || record.administration || "Sem ato destacado")}</small>
          </div>
        </div>

        <div class="record-meta">
          <div class="meta-block">
            <span>Fornecedor</span>
            <strong>${escapeHtml(record.supplier || "Não identificado")}</strong>
            <small>${escapeHtml(record.valueLabel || formatCurrency(record.valueNumber))}</small>
          </div>
          <div class="meta-block">
            <span>Período de gestão</span>
            <strong>${escapeHtml(record.administration || "Não identificado")}</strong>
            <small>${escapeHtml(record.year ? `Ano de referência ${record.year}` : "Ano de referência não identificado")}</small>
          </div>
          <div class="meta-block">
            <span>Cruzamento</span>
            <strong>${escapeHtml(getLabel("source", record.sourceStatus))}</strong>
            <small>${escapeHtml(`${record.movementCount || 0} movimentação(ões) associadas`)}</small>
          </div>
          <div class="meta-block">
            <span>Prazo final</span>
            <strong>${record.vigency?.endDate ? escapeHtml(formatDate(record.vigency.endDate)) : "Sem data final"}</strong>
            <small>${record.vigency?.daysUntilEnd != null ? `${escapeHtml(String(record.vigency.daysUntilEnd))} dia(s) estimados` : "Sem prazo consolidado"}</small>
          </div>
        </div>

        ${renderAlertPills(record)}

        <div class="record-actions">
          ${diaryLink}
          ${portalLink}
        </div>
      </article>
    `;
  }

  function renderAlertRecords() {
    const records = getCurrentRecords()
      .filter((record) => getUniqueAlerts(record.alerts).length > 0)
      .slice(0, 10);

    if (!records.length) {
      elements.alertRecords.innerHTML = `<div class="empty-state">Nenhum alerta prioritário foi encontrado.</div>`;
      return;
    }

    elements.alertRecords.innerHTML = records.map(renderRecordCard).join("");
  }

  function renderResults() {
    const filtered = getFilteredRecords();
    const visible = filtered.slice(0, state.visibleCount);
    const currentInScope =
      state.filters.scope === "atuais"
        ? filtered.length
        : filtered.filter((record) => record.vigency?.isCurrent).length;

    elements.resultsMeta.textContent =
      `${formatNumber(filtered.length)} contrato(s) exibidos. ` +
      (state.filters.scope === "atuais"
        ? "A consulta está limitada aos contratos tratados como atuais."
        : `${formatNumber(currentInScope)} deles ainda aparecem como atuais na leitura cruzada.`);

    const activeFilters = [];
    if (state.filters.query) activeFilters.push(`Busca: ${state.filters.query}`);
    if (state.filters.organization) activeFilters.push(`Órgão: ${state.filters.organization}`);
    if (state.filters.administration) activeFilters.push(`Gestão: ${state.filters.administration}`);
    if (state.filters.vigency !== "todos") activeFilters.push(`Vigência: ${getLabel("vigency", state.filters.vigency)}`);
    if (state.filters.management !== "todos") activeFilters.push(`Responsável: ${getLabel("management", state.filters.management)}`);
    if (state.filters.source !== "todos") activeFilters.push(`Fonte: ${getLabel("source", state.filters.source)}`);
    if (state.filters.scope !== DEFAULT_FILTERS.scope) activeFilters.push(`Escopo: ${getLabel("scope", state.filters.scope)}`);
    elements.activeFilterSummary.textContent = activeFilters.length
      ? activeFilters.join(" · ")
      : "Sem filtros adicionais além do escopo padrão de contratos atuais.";

    if (!visible.length) {
      elements.recordList.innerHTML = `<div class="empty-state">Nenhum contrato encontrou correspondência com os filtros aplicados.</div>`;
      elements.loadMore.classList.add("hidden");
      return;
    }

    elements.recordList.innerHTML = visible.map(renderRecordCard).join("");
    elements.loadMore.classList.toggle("hidden", visible.length >= filtered.length);
  }

  function renderFooter() {
    const summary = state.payload.summary;
    elements.footerCopy.textContent =
      `Base com ${formatNumber(summary.analisados)} edições analisadas e ${formatNumber(summary.contratosPortal)} contratos do portal considerados no cruzamento público.`;
  }

  function renderAll() {
    if (!state.payload) return;
    renderHero();
    renderSummaryCards();
    renderMethodology();
    renderStatusGrid();
    renderPriorityGroups();
    renderOrganizationSummary();
    renderQuickPresets();
    renderAlertRecords();
    renderResults();
    renderFooter();
    syncUrlState();
  }

  function resetFilters() {
    state.filters = { ...DEFAULT_FILTERS };
    state.visibleCount = PAGE_SIZE;
    syncControls();
    syncUrlState();
    renderAll();
  }

  function applyPreset(presetKey) {
    const preset = PRESETS[presetKey];
    if (!preset) return;
    state.filters = { ...DEFAULT_FILTERS, ...preset.filters };
    state.visibleCount = PAGE_SIZE;
    syncControls();
    setView("contracts");
    renderAll();
  }

  function bindEvents() {
    elements.viewButtons.forEach((button) => {
      button.addEventListener("click", () => {
        setView(button.dataset.viewButton);
      });
    });

    elements.searchInput.addEventListener("input", () => {
      state.filters.query = elements.searchInput.value.trim();
      state.visibleCount = PAGE_SIZE;
      renderAll();
    });

    [
      [elements.scopeSelect, "scope"],
      [elements.organizationSelect, "organization"],
      [elements.administrationSelect, "administration"],
      [elements.vigencySelect, "vigency"],
      [elements.managementSelect, "management"],
      [elements.sourceSelect, "source"],
    ].forEach(([element, key]) => {
      element.addEventListener("change", () => {
        state.filters[key] = element.value;
        state.visibleCount = PAGE_SIZE;
        renderAll();
      });
    });

    elements.clearFilters.addEventListener("click", () => {
      resetFilters();
    });

    elements.loadMore.addEventListener("click", () => {
      state.visibleCount += PAGE_SIZE;
      renderResults();
    });

    document.addEventListener("click", (event) => {
      const presetButton = event.target.closest("[data-preset]");
      if (!presetButton) return;
      applyPreset(presetButton.dataset.preset);
    });
  }

  async function bootstrap() {
    const initialState = readUrlState();
    state.view = initialState.view;
    state.filters = { ...DEFAULT_FILTERS, ...initialState.filters };

    const response = await fetch("./data/contracts-dashboard.json", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Falha ao carregar a base pública (${response.status}).`);
    }

    state.payload = await response.json();

    populateSelect(elements.scopeSelect, ["atuais", "todos"], "scope", "atuais");
    populateSelect(elements.vigencySelect, (state.payload.filters.vigencyStates || []).filter((value) => value !== "todos"), "vigency");
    populateSelect(elements.managementSelect, (state.payload.filters.managementStates || []).filter((value) => value !== "todos"), "management");
    populateSelect(elements.sourceSelect, (state.payload.filters.sourceStates || []).filter((value) => value !== "todos"), "source");

    const setPlainOptions = (select, values, emptyLabel) => {
      const previous = select.value;
      select.innerHTML = "";
      const empty = document.createElement("option");
      empty.value = "";
      empty.textContent = emptyLabel;
      select.appendChild(empty);
      values.forEach((value) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        select.appendChild(option);
      });
      if ([...select.options].some((option) => option.value === previous)) {
        select.value = previous;
      }
    };

    setPlainOptions(elements.organizationSelect, state.payload.filters.organizations || [], "Todos os órgãos");
    setPlainOptions(elements.administrationSelect, state.payload.filters.administrations || [], "Todas as gestões");
    syncControls();
    setView(state.view);
    bindEvents();
    renderAll();
  }

  bootstrap().catch((error) => {
    const message = escapeHtml(error.message || "Não foi possível carregar a base pública.");
    elements.heroSummary.textContent = "Falha ao carregar os dados públicos do painel.";
    elements.heroCallout.textContent = error.message || "Não foi possível carregar a base pública.";
    elements.summaryCards.innerHTML = `<div class="empty-state">${message}</div>`;
    elements.methodNotes.innerHTML = `<div class="empty-state">${message}</div>`;
    elements.statusGrid.innerHTML = `<div class="empty-state">${message}</div>`;
    elements.priorityGroups.innerHTML = `<div class="empty-state">${message}</div>`;
    elements.organizationSummary.innerHTML = `<div class="empty-state">${message}</div>`;
    elements.alertRecords.innerHTML = `<div class="empty-state">${message}</div>`;
    elements.recordList.innerHTML = `<div class="empty-state">${message}</div>`;
    elements.loadMore.classList.add("hidden");
  });
})();
