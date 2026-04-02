(() => {
  const PAGE_SIZE = 18;
  const SEARCH_DEBOUNCE_MS = 140;
  const MAX_QUERY_LENGTH = 160;
  const FETCH_TIMEOUT_MS = 15000;
  const DEFAULT_VIEW = "overview";
  const VALID_VIEWS = new Set(["overview", "alerts", "contracts"]);
  const ALLOWED_EXTERNAL_HOSTS = new Set(["iguape.sp.gov.br", "www.iguape.sp.gov.br"]);
  const collator = new Intl.Collator("pt-BR", { numeric: true, sensitivity: "base" });
  const trustedTypesPolicy = window.trustedTypes?.createPolicy("contracts-dashboard", {
    createHTML: (value) => String(value),
  });

  const DEFAULT_FILTERS = {
    query: "",
    organization: "",
    administration: "",
    vigency: "todos",
    management: "todos",
    source: "todos",
    criticality: "todos",
    sort: "prioridade",
    scope: "atuais",
  };

  const LABELS = {
    vigency: {
      todos: "Todas as situações",
      vigente_confirmado: "Vigência confirmada",
      vigente_inferido: "Vigência estimada",
      em_acompanhamento: "Em validação",
      encerrado: "Encerrado",
      sem_sinal_atual: "Sem evidência recente",
    },
    management: {
      todos: "Todas as situações",
      completos: "Designações completas",
      sem_gestor: "Sem gestor designado",
      sem_fiscal: "Sem fiscal designado",
      sem_gestor_e_fiscal: "Sem responsáveis designados",
      revisao: "Em revisão",
      exoneracao: "Indício de exoneração",
    },
    source: {
      todos: "Todas as fontes",
      cruzado: "Cruzamento confirmado",
      somente_diario: "Apenas Diário Oficial",
      somente_portal: "Apenas Portal",
    },
    criticality: {
      todos: "Todas as criticidades",
      alta: "Alta criticidade",
      media: "Média criticidade",
      baixa: "Baixa criticidade",
    },
    sort: {
      prioridade: "Prioridade operacional",
      movimentacao_recente: "Movimentação mais recente",
      prazo_mais_proximo: "Prazo mais próximo",
      maior_valor: "Maior valor",
      orgao: "Órgão",
      contrato: "Número do contrato",
    },
    scope: {
      atuais: "Apenas vigentes",
      todos: "Todos os registros",
    },
  };

  const PRESETS = {
    semGestorEFiscal: {
      label: "Sem responsáveis",
      description: "Registros sem gestor e fiscal.",
      filters: { management: "sem_gestor_e_fiscal", scope: "atuais", sort: "prioridade" },
    },
    semFiscal: {
      label: "Sem fiscal",
      description: "Registros sem fiscal atual.",
      filters: { management: "sem_fiscal", scope: "atuais", sort: "prioridade" },
    },
    altaCriticidade: {
      label: "Alta criticidade",
      description: "Pendências com maior peso operacional.",
      filters: { criticality: "alta", scope: "atuais", sort: "prioridade" },
    },
    somenteDiario: {
      label: "Apenas Diário",
      description: "Registros sem confirmação cruzada.",
      filters: { source: "somente_diario", scope: "atuais", sort: "prioridade" },
    },
    completos: {
      label: "Designações completas",
      description: "Gestor e fiscal identificados.",
      filters: { management: "completos", scope: "atuais", sort: "movimentacao_recente" },
    },
  };

  const state = {
    payload: null,
    view: DEFAULT_VIEW,
    filters: { ...DEFAULT_FILTERS },
    visibleCount: PAGE_SIZE,
    filteredCache: new Map(),
    searchTimer: null,
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
    coverageGrid: document.getElementById("coverage-grid"),
    sourceGrid: document.getElementById("source-grid"),
    deadlineGrid: document.getElementById("deadline-grid"),
    pendingDeadlineSummary: document.getElementById("pending-deadline-summary"),
    pendingDeadlineRecords: document.getElementById("pending-deadline-records"),
    recentMovements: document.getElementById("recent-movements"),
    insightList: document.getElementById("insight-list"),
    alertsSummary: document.getElementById("alerts-summary"),
    alertSummaryGrid: document.getElementById("alert-summary-grid"),
    alertRecords: document.getElementById("alert-records"),
    reviewQueueSummary: document.getElementById("review-queue-summary"),
    reviewSummaryGrid: document.getElementById("review-summary-grid"),
    reviewQueue: document.getElementById("review-queue"),
    searchInput: document.getElementById("search-input"),
    scopeSelect: document.getElementById("scope-select"),
    organizationSelect: document.getElementById("organization-select"),
    administrationSelect: document.getElementById("administration-select"),
    vigencySelect: document.getElementById("vigency-select"),
    managementSelect: document.getElementById("management-select"),
    sourceSelect: document.getElementById("source-select"),
    criticalitySelect: document.getElementById("criticality-select"),
    sortSelect: document.getElementById("sort-select"),
    quickPresets: document.getElementById("quick-presets"),
    resultsMeta: document.getElementById("results-meta"),
    activeFilterSummary: document.getElementById("active-filter-summary"),
    resultInsightGrid: document.getElementById("result-insight-grid"),
    recordList: document.getElementById("record-list"),
    loadMore: document.getElementById("load-more"),
    clearFilters: document.getElementById("clear-filters"),
    shareView: document.getElementById("share-view"),
    exportCsv: document.getElementById("export-csv"),
    footerCopy: document.getElementById("footer-copy"),
    viewButtons: [...document.querySelectorAll("[data-view-button]")],
    viewPanels: [...document.querySelectorAll("[data-view-panel]")],
  };

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function setHtml(element, html) {
    const value = String(html ?? "");
    element.innerHTML = trustedTypesPolicy ? trustedTypesPolicy.createHTML(value) : value;
  }

  function sanitizePlainText(value, limit = 400) {
    return repairText(value)
      .replace(/[\u0000-\u001f\u007f]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, limit);
  }

  function sanitizeQueryInput(value) {
    return sanitizePlainText(value, MAX_QUERY_LENGTH);
  }

  function sanitizeExternalUrl(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";

    try {
      const url = new URL(raw, window.location.origin);
      if (url.origin === window.location.origin) {
        return `${url.pathname}${url.search}${url.hash}`;
      }
      if (url.protocol !== "https:") return "";
      if (!ALLOWED_EXTERNAL_HOSTS.has(url.hostname.toLowerCase())) return "";
      return url.toString();
    } catch {
      return "";
    }
  }

  function validatePayloadShape(payload) {
    if (!payload || typeof payload !== "object") return false;
    if (!payload.summary || typeof payload.summary !== "object") return false;
    if (!payload.reviewSummary || typeof payload.reviewSummary !== "object") return false;
    if (!Array.isArray(payload.records)) return false;
    if (!Array.isArray(payload.reviewQueue)) return false;
    if (!payload.filters || typeof payload.filters !== "object") return false;
    return true;
  }

  function qualityScore(text) {
    const printable = (text.match(/[0-9A-Za-zÀ-ÿ\s.,;:!?()/%ºª°"'/\-]/g) || []).length;
    const noise = (text.match(/[ÃÂ�┬]/g) || []).length;
    return printable - noise * 3;
  }

  function repairText(value) {
    const original = String(value ?? "").replace(/\s+/g, " ").trim();
    if (!original || !/[ÃÂ�┬]/.test(original)) return original;
    if (typeof TextDecoder === "undefined") return original;

    try {
      const bytes = Uint8Array.from([...original].map((character) => character.charCodeAt(0) & 0xff));
      const decoded = new TextDecoder("utf-8", { fatal: false }).decode(bytes).replace(/\u0000/g, "").trim();
      if (decoded && qualityScore(decoded) > qualityScore(original)) {
        return decoded;
      }
    } catch {
      return original;
    }

    return original;
  }

  function normalizeText(value) {
    return repairText(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function parseDate(value) {
    if (!value) return null;
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return null;
    return date;
  }

  function formatDate(value) {
    const date = value instanceof Date ? value : parseDate(value);
    if (!date) return "Não informado";
    return new Intl.DateTimeFormat("pt-BR", { dateStyle: "medium" }).format(date);
  }

  function formatDateTime(value) {
    const date = value instanceof Date ? value : parseDate(value);
    if (!date) return "Não informado";
    return new Intl.DateTimeFormat("pt-BR", { dateStyle: "medium", timeStyle: "short" }).format(date);
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("pt-BR").format(Number(value || 0));
  }

  function formatPercent(value, total) {
    if (!total) return "0%";
    const percentage = (Number(value || 0) / Number(total || 1)) * 100;
    return `${new Intl.NumberFormat("pt-BR", { maximumFractionDigits: percentage >= 10 ? 0 : 1 }).format(percentage)}%`;
  }

  function formatCurrency(value) {
    const numeric = Number(value || 0);
    if (!numeric || numeric <= 0) return "Sem valor";
    return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(numeric);
  }

  function truncateText(value, limit = 260) {
    const text = repairText(value).replace(/\s+/g, " ").trim();
    if (text.length <= limit) return text;
    return `${text.slice(0, limit - 3).trim()}...`;
  }

  function formatDayDelta(days) {
    if (days == null || Number.isNaN(Number(days))) return "Sem prazo definido";
    const numeric = Number(days);
    if (numeric < 0) return `Expirado há ${formatNumber(Math.abs(numeric))} dia(s)`;
    if (numeric === 0) return "Vence hoje";
    return `${formatNumber(numeric)} dia(s) restantes`;
  }

  function getLabel(group, value) {
    return LABELS[group]?.[value] || repairText(value) || "Não informado";
  }

  function getCriticalityWeight(level) {
    if (level === "alta") return 3;
    if (level === "media") return 2;
    return 1;
  }

  function getManagementPriority(value) {
    if (value === "sem_gestor_e_fiscal") return 4;
    if (value === "sem_gestor") return 3;
    if (value === "sem_fiscal") return 2;
    if (value === "revisao" || value === "exoneracao") return 1;
    return 0;
  }

  function getSourcePriority(value) {
    if (value === "somente_diario") return 3;
    if (value === "somente_portal") return 2;
    if (value === "cruzado") return 1;
    return 0;
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

  function deriveCriticality(record) {
    if ((record.alertWeight || 0) >= 3 || record.managementState === "sem_gestor_e_fiscal") return "alta";
    if (
      (record.alertWeight || 0) >= 2 ||
      record.managementState === "sem_gestor" ||
      record.managementState === "sem_fiscal" ||
      record.managementState === "revisao" ||
      record.managementState === "exoneracao"
    ) {
      return "media";
    }
    return "baixa";
  }

  function preprocessPayload(payload) {
    const generatedAt = parseDate(payload.generatedAt);
    const records = (payload.records || []).map((record, index) => preprocessRecord(record, index));

    return {
      ...payload,
      methodology: {
        title: sanitizePlainText(payload.methodology?.title),
        summary: sanitizePlainText(payload.methodology?.summary, 800),
        notes: (payload.methodology?.notes || []).map((note) => sanitizePlainText(note, 300)),
      },
      filters: {
        ...payload.filters,
        organizations: (payload.filters?.organizations || []).map((value) => sanitizePlainText(value, 160)),
        administrations: (payload.filters?.administrations || []).map((value) => sanitizePlainText(value, 160)),
      },
      organizationSummary: (payload.organizationSummary || []).map((item) => ({
        organization: sanitizePlainText(item.organization, 160),
        count: Number(item.count || 0),
      })),
      reviewSummary: {
        total: Number(payload.reviewSummary?.total || 0),
        current: Number(payload.reviewSummary?.current || 0),
        high: Number(payload.reviewSummary?.high || 0),
        medium: Number(payload.reviewSummary?.medium || 0),
        low: Number(payload.reviewSummary?.low || 0),
        divergent: Number(payload.reviewSummary?.divergent || 0),
        crossPending: Number(payload.reviewSummary?.crossPending || 0),
        operationalLow: Number(payload.reviewSummary?.operationalLow || 0),
        documentalLow: Number(payload.reviewSummary?.documentalLow || 0),
      },
      reviewQueue: (payload.reviewQueue || []).map((item, index) => ({
        id: sanitizePlainText(item.id || `review-${index}`, 120),
        masterContractId: sanitizePlainText(item.masterContractId, 120),
        normalizedKey: sanitizePlainText(item.normalizedKey, 80),
        contractNumber: sanitizePlainText(item.contractNumber, 80),
        administration: sanitizePlainText(item.administration, 120),
        organization: sanitizePlainText(item.organization, 180),
        sourceStatus: sanitizePlainText(item.sourceStatus, 60),
        managementState: sanitizePlainText(item.managementState, 60),
        isCurrent: Boolean(item.isCurrent),
        priority: sanitizePlainText(item.priority, 20),
        priorityWeight: Number(item.priorityWeight || 0),
        sourceAlignment: sanitizePlainText(item.sourceAlignment, 40),
        reasonSummary: sanitizePlainText(item.reasonSummary, 280),
        reasonCount: Number(item.reasonCount || 0),
        candidateCount: Number(item.candidateCount || 0),
        recommendedConfidence: sanitizePlainText(item.recommendedConfidence, 20),
        recommendedScore: Number(item.recommendedScore || 0),
        divergenceCount: Number(item.divergenceCount || 0),
        divergenceTypes: (item.divergenceTypes || []).map((value) => sanitizePlainText(value, 60)),
        criticalMissingFields: (item.criticalMissingFields || []).map((value) => sanitizePlainText(value, 60)),
        overallConfidence: sanitizePlainText(item.overallConfidence, 20),
        operationalConfidence: sanitizePlainText(item.operationalConfidence, 20),
        documentalConfidence: sanitizePlainText(item.documentalConfidence, 20),
        publishedAt: parseDate(item.publishedAt),
        managementActAt: parseDate(item.managementActAt),
        endDate: parseDate(item.endDate),
        daysUntilEnd:
          item.daysUntilEnd == null || Number.isNaN(Number(item.daysUntilEnd))
            ? null
            : Number(item.daysUntilEnd),
        links: {
          diary: sanitizeExternalUrl(item.links?.diary),
          portal: sanitizeExternalUrl(item.links?.portal),
        },
      })),
      generatedAt,
      records,
    };
  }

  function preprocessRecord(record, index) {
    const manager = {
      ...record.manager,
      name: sanitizePlainText(record.manager?.name, 160),
      role: sanitizePlainText(record.manager?.role, 220),
    };

    const inspector = {
      ...record.inspector,
      name: sanitizePlainText(record.inspector?.name, 160),
      role: sanitizePlainText(record.inspector?.role, 220),
    };

    const alerts = (record.alerts || []).map((alert) => ({
      ...alert,
      title: sanitizePlainText(alert.title, 220),
      description: sanitizePlainText(alert.description, 500),
    }));

    const vigency = {
      ...record.vigency,
      label: sanitizePlainText(record.vigency?.label, 220),
      sourceLabel: sanitizePlainText(record.vigency?.sourceLabel, 220),
      daysUntilEnd:
        record.vigency?.daysUntilEnd == null || Number.isNaN(Number(record.vigency?.daysUntilEnd))
          ? null
          : Number(record.vigency.daysUntilEnd),
    };

    const movementDate = parseDate(record.managementActAt || record.publishedAt);
    const endDate = parseDate(vigency.endDate);
    const criticality = deriveCriticality({ ...record, alerts });

    return {
      ...record,
      id: record.id || `record-${index}`,
      contractNumber: sanitizePlainText(record.contractNumber, 80),
      processNumber: sanitizePlainText(record.processNumber, 120),
      administration: sanitizePlainText(record.administration, 160),
      organization: sanitizePlainText(record.organization, 160),
      supplier: sanitizePlainText(record.supplier, 240),
      object: sanitizePlainText(record.object, 1200),
      valueLabel: sanitizePlainText(record.valueLabel, 120),
      managementSummary: sanitizePlainText(record.managementSummary, 300),
      lastMovementTitle: sanitizePlainText(record.lastMovementTitle, 220),
      normalizedKey: sanitizePlainText(record.normalizedKey, 80),
      managerPersonnelStatus: sanitizePlainText(record.managerPersonnelStatus, 40),
      inspectorPersonnelStatus: sanitizePlainText(record.inspectorPersonnelStatus, 40),
      managerExonerationSignal: Boolean(record.managerExonerationSignal),
      inspectorExonerationSignal: Boolean(record.inspectorExonerationSignal),
      lifecycle: {
        ...record.lifecycle,
        summary: sanitizePlainText(record.lifecycle?.summary, 220),
        eventCount: Number(record.lifecycle?.eventCount || 0),
        additiveCount: Number(record.lifecycle?.additiveCount || 0),
        apostilleCount: Number(record.lifecycle?.apostilleCount || 0),
        terminationCount: Number(record.lifecycle?.terminationCount || 0),
        isAdditivado: Boolean(record.lifecycle?.isAdditivado),
        hasActiveTermination: Boolean(record.lifecycle?.hasActiveTermination),
      },
      additives: {
        ...record.additives,
        isAdditivado: Boolean(record.additives?.isAdditivado),
        totalKnown: Number(record.additives?.totalKnown || 0),
        portalCount: Number(record.additives?.portalCount || 0),
        diaryCount: Number(record.additives?.diaryCount || 0),
        apostilleCount: Number(record.additives?.apostilleCount || 0),
        terminationCount: Number(record.additives?.terminationCount || 0),
        hasActiveTermination: Boolean(record.additives?.hasActiveTermination),
      },
      review: {
        ...record.review,
        required: Boolean(record.review?.required),
        priority: sanitizePlainText(record.review?.priority, 24),
        sourceAlignment: sanitizePlainText(record.review?.sourceAlignment, 32),
        reasonSummary: sanitizePlainText(record.review?.reasonSummary, 220),
        reasonCount: Number(record.review?.reasonCount || 0),
        candidateCount: Number(record.review?.candidateCount || 0),
        divergenceCount: Number(record.review?.divergenceCount || 0),
      },
      manager,
      inspector,
      alerts,
      vigency,
      links: {
        diary: sanitizeExternalUrl(record.links?.diary),
        portal: sanitizeExternalUrl(record.links?.portal),
      },
      _movementDate: movementDate,
      _movementTimestamp: movementDate ? movementDate.getTime() : 0,
      _endDate: endDate,
      _endTimestamp: endDate ? endDate.getTime() : Number.MAX_SAFE_INTEGER,
      _criticality: criticality,
      _criticalityWeight: getCriticalityWeight(criticality),
      _hasManager: Boolean(manager.name),
      _hasInspector: Boolean(inspector.name),
      _hasCompleteAssignments: Boolean(manager.name && inspector.name),
      _isCurrent: Boolean(vigency.isCurrent),
      _valueNumber: Number(record.valueNumber || 0),
      _searchText: normalizeText(
        [
          record.contractNumber,
          record.processNumber,
          record.normalizedKey,
          record.organization,
          record.supplier,
          record.object,
          record.managementSummary,
          manager.name,
          manager.role,
          inspector.name,
          inspector.role,
          record.administration,
          record.year,
          record.lastMovementTitle,
          ...alerts.map((alert) => alert.title),
        ].join(" ")
      ),
    };
  }

  function readUrlState() {
    const params = new URLSearchParams(window.location.search);
    return {
      view: VALID_VIEWS.has(params.get("view")) ? params.get("view") : DEFAULT_VIEW,
      filters: {
        query: sanitizeQueryInput(params.get("q") || ""),
        organization: sanitizePlainText(params.get("org") || "", 160),
        administration: sanitizePlainText(params.get("adm") || "", 160),
        vigency: params.get("vig") || DEFAULT_FILTERS.vigency,
        management: params.get("mgmt") || DEFAULT_FILTERS.management,
        source: params.get("src") || DEFAULT_FILTERS.source,
        criticality: params.get("crit") || DEFAULT_FILTERS.criticality,
        sort: params.get("sort") || DEFAULT_FILTERS.sort,
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
    if (state.filters.criticality !== DEFAULT_FILTERS.criticality) params.set("crit", state.filters.criticality);
    if (state.filters.sort !== DEFAULT_FILTERS.sort) params.set("sort", state.filters.sort);
    if (state.filters.scope !== DEFAULT_FILTERS.scope) params.set("scope", state.filters.scope);

    const query = params.toString();
    const nextUrl = query ? `${window.location.pathname}?${query}` : window.location.pathname;
    window.history.replaceState({}, "", nextUrl);
  }

  function getRecords() {
    return state.payload?.records || [];
  }

  function getCurrentRecords() {
    return getRecords().filter((record) => record._isCurrent);
  }

  function summarizeRecordSet(records) {
    const total = records.length;
    const withCompleteAssignments = records.filter((record) => record.managementState === "completos").length;
    const withoutManager = records.filter((record) => record.managementState === "sem_gestor").length;
    const withoutInspector = records.filter((record) => record.managementState === "sem_fiscal").length;
    const withoutBoth = records.filter((record) => record.managementState === "sem_gestor_e_fiscal").length;
    const critical = records.filter((record) => record._criticality === "alta").length;
    const crossed = records.filter((record) => record.sourceStatus === "cruzado").length;
    const diaryOnly = records.filter((record) => record.sourceStatus === "somente_diario").length;
    const portalOnly = records.filter((record) => record.sourceStatus === "somente_portal").length;
    const withDeadline = records.filter((record) => record._endDate).length;
    const expiringSoon = records.filter((record) => record.vigency?.daysUntilEnd != null && record.vigency.daysUntilEnd >= 0 && record.vigency.daysUntilEnd <= 30).length;
    const recentThreshold = (state.payload?.generatedAt?.getTime() || Date.now()) - 30 * 24 * 60 * 60 * 1000;
    const recentMovements = records.filter((record) => record._movementTimestamp >= recentThreshold).length;

    return {
      total,
      withCompleteAssignments,
      withoutManager,
      withoutInspector,
      withoutBoth,
      critical,
      crossed,
      diaryOnly,
      portalOnly,
      withDeadline,
      expiringSoon,
      recentMovements,
    };
  }

  function getTopOrganization(records) {
    const counts = new Map();
    records.forEach((record) => {
      const key = record.organization || "Não informado";
      counts.set(key, (counts.get(key) || 0) + 1);
    });

    let top = null;
    counts.forEach((count, organization) => {
      if (!top || count > top.count) top = { organization, count };
    });

    return top;
  }

  function compareByPriority(a, b) {
    if (b._criticalityWeight !== a._criticalityWeight) return b._criticalityWeight - a._criticalityWeight;
    if (getManagementPriority(b.managementState) !== getManagementPriority(a.managementState)) {
      return getManagementPriority(b.managementState) - getManagementPriority(a.managementState);
    }
    if (getSourcePriority(b.sourceStatus) !== getSourcePriority(a.sourceStatus)) {
      return getSourcePriority(b.sourceStatus) - getSourcePriority(a.sourceStatus);
    }
    if (b._movementTimestamp !== a._movementTimestamp) return b._movementTimestamp - a._movementTimestamp;
    if ((a.vigency?.daysUntilEnd ?? Number.MAX_SAFE_INTEGER) !== (b.vigency?.daysUntilEnd ?? Number.MAX_SAFE_INTEGER)) {
      return (a.vigency?.daysUntilEnd ?? Number.MAX_SAFE_INTEGER) - (b.vigency?.daysUntilEnd ?? Number.MAX_SAFE_INTEGER);
    }
    return collator.compare(a.contractNumber || "", b.contractNumber || "");
  }

  function compareRecords(a, b) {
    switch (state.filters.sort) {
      case "movimentacao_recente":
        if (b._movementTimestamp !== a._movementTimestamp) return b._movementTimestamp - a._movementTimestamp;
        return compareByPriority(a, b);
      case "prazo_mais_proximo": {
        const aDays = a.vigency?.daysUntilEnd ?? Number.MAX_SAFE_INTEGER;
        const bDays = b.vigency?.daysUntilEnd ?? Number.MAX_SAFE_INTEGER;
        if (aDays !== bDays) return aDays - bDays;
        return compareByPriority(a, b);
      }
      case "maior_valor":
        if (b._valueNumber !== a._valueNumber) return b._valueNumber - a._valueNumber;
        return compareByPriority(a, b);
      case "orgao": {
        const orgCompare = collator.compare(a.organization || "", b.organization || "");
        if (orgCompare !== 0) return orgCompare;
        return collator.compare(a.contractNumber || "", b.contractNumber || "");
      }
      case "contrato": {
        const yearDiff = Number(b.year || 0) - Number(a.year || 0);
        if (yearDiff !== 0) return yearDiff;
        return collator.compare(a.contractNumber || "", b.contractNumber || "");
      }
      default:
        return compareByPriority(a, b);
    }
  }

  function matchesQuery(record, query) {
    if (!query) return true;
    return record._searchText.includes(normalizeText(query));
  }

  function getCacheKey() {
    return JSON.stringify(state.filters);
  }

  function getFilteredRecords() {
    const cacheKey = getCacheKey();
    if (state.filteredCache.has(cacheKey)) {
      return state.filteredCache.get(cacheKey);
    }

    const filtered = getRecords()
      .filter((record) => {
        if (state.filters.scope === "atuais" && !record._isCurrent) return false;
        if (state.filters.organization && record.organization !== state.filters.organization) return false;
        if (state.filters.administration && record.administration !== state.filters.administration) return false;
        if (state.filters.vigency !== "todos" && record.vigency?.state !== state.filters.vigency) return false;
        if (state.filters.management !== "todos" && record.managementState !== state.filters.management) return false;
        if (state.filters.source !== "todos" && record.sourceStatus !== state.filters.source) return false;
        if (state.filters.criticality !== "todos" && record._criticality !== state.filters.criticality) return false;
        if (!matchesQuery(record, state.filters.query)) return false;
        return true;
      })
      .sort(compareRecords);

    state.filteredCache.set(cacheKey, filtered);
    return filtered;
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
    elements.criticalitySelect.value = state.filters.criticality;
    elements.sortSelect.value = state.filters.sort;
  }

  function setLabeledOptions(select, values, groupName) {
    const previous = select.value;
    setHtml(select, "");

    values.forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = getLabel(groupName, value);
      select.appendChild(option);
    });

    if ([...select.options].some((option) => option.value === previous)) {
      select.value = previous;
    }
  }

  function setPlainOptions(select, values, emptyLabel) {
    const previous = select.value;
    setHtml(select, "");

    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = emptyLabel;
    select.appendChild(emptyOption);

    values.forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });

    if ([...select.options].some((option) => option.value === previous)) {
      select.value = previous;
    }
  }

  function ensureValidFilterValue(key, select) {
    if ([...select.options].some((option) => option.value === state.filters[key])) return;
    state.filters[key] = DEFAULT_FILTERS[key];
  }

  function renderCardCollection(container, cards) {
    setHtml(
      container,
      cards
      .map(
        (card) => `
          <article class="analysis-card ${card.className || ""}">
            <span class="eyebrow">${escapeHtml(card.label)}</span>
            <strong class="analysis-value">${escapeHtml(card.value)}</strong>
            <span class="analysis-detail">${escapeHtml(card.detail)}</span>
          </article>
        `
      )
      .join("")
    );
  }

  function renderSummaryCards() {
    const summary = state.payload.summary;
    const currentTotal = summary.contratosAtuais || 0;
    const reviewCurrent = state.payload.reviewSummary?.current || 0;
    const cards = [
      {
        label: "Contratos vigentes",
        value: formatNumber(currentTotal),
        detail: "Registros atuais monitorados",
      },
      {
        label: "Designações completas",
        value: formatNumber(summary.comResponsaveisCompletos),
        detail: formatPercent(summary.comResponsaveisCompletos, currentTotal),
      },
      {
        label: "Apenas Diário Oficial",
        value: formatNumber(summary.somenteDiario),
        detail: "Sem confirmação cruzada",
      },
      {
        label: "Ocorrências críticas",
        value: formatNumber(summary.alertasCriticos),
        detail: "Demandam providência",
      },
    ];

    cards[2] = {
      label: "RevisÃ£o dirigida",
      value: formatNumber(reviewCurrent),
      detail: "Casos com conferÃªncia focal",
    };

    setHtml(
      elements.summaryCards,
      cards
      .map(
        (card) => `
          <article class="metric-card">
            <span class="eyebrow">${escapeHtml(card.label)}</span>
            <strong class="metric-value">${escapeHtml(card.value)}</strong>
            <span class="metric-detail">${escapeHtml(card.detail)}</span>
          </article>
        `
      )
      .join("")
    );
  }

  function renderSummaryCardsV2() {
    const summary = state.payload.summary;
    const currentTotal = summary.contratosAtuais || 0;
    const reviewCurrent = state.payload.reviewSummary?.current || 0;
    const cards = [
      {
        label: "Contratos vigentes",
        value: formatNumber(currentTotal),
        detail: "Registros atuais monitorados",
      },
      {
        label: "Designacoes completas",
        value: formatNumber(summary.comResponsaveisCompletos),
        detail: formatPercent(summary.comResponsaveisCompletos, currentTotal),
      },
      {
        label: "Revisao dirigida",
        value: formatNumber(reviewCurrent),
        detail: "Casos com conferencia focal",
      },
      {
        label: "Ocorrencias criticas",
        value: formatNumber(summary.alertasCriticos),
        detail: "Demandam providencia",
      },
    ];

    setHtml(
      elements.summaryCards,
      cards
        .map(
          (card) => `
            <article class="metric-card">
              <span class="eyebrow">${escapeHtml(card.label)}</span>
              <strong class="metric-value">${escapeHtml(card.value)}</strong>
              <span class="metric-detail">${escapeHtml(card.detail)}</span>
            </article>
          `
        )
        .join("")
    );
  }

  function renderMethodology() {
    const summary = state.payload.summary || {};
    const currentTotal = summary.contratosAtuais || 0;
    const alertsTotal = summary.alertasCriticos || 0;
    const searchableTotal = state.payload.records?.length || 0;

    elements.updatedAt.textContent = `Atualizado em ${formatDateTime(state.payload.generatedAt)}`;
    elements.methodSummary.textContent = "Escolha a \u00E1rea e siga pela navega\u00E7\u00E3o do painel.";
    setHtml(
      elements.methodNotes,
      [
        {
          label: "Painel",
          title: "Vis\u00E3o geral",
          detail: `${formatNumber(currentTotal)} contratos vigentes monitorados`,
          view: "overview",
        },
        {
          label: "Ocorr\u00EAncias",
          title: "Prioridades",
          detail: `${formatNumber(alertsTotal)} registros com criticidade alta`,
          view: "alerts",
        },
        {
          label: "Consulta",
          title: "Pesquisa",
          detail: `${formatNumber(searchableTotal)} registros dispon\u00EDveis para busca`,
          view: "contracts",
        },
      ]
        .map(
          (item) => `
            <article class="note-card note-card--module">
              <span class="eyebrow">${escapeHtml(item.label)}</span>
              <strong>${escapeHtml(item.title)}</strong>
              <span>${escapeHtml(item.detail)}</span>
              <button class="secondary-button secondary-button--inline" type="button" data-open-view="${escapeHtml(item.view)}">
                Abrir
              </button>
            </article>
          `
        )
        .join("")
    );
  }

  function renderHero() {
    const summary = state.payload.summary;
    const currentTotal = summary.contratosAtuais || 0;
    const withoutFiscal = summary.semFiscal || 0;
    const topOrganization = getTopOrganization(getCurrentRecords());

    elements.heroSummary.textContent = `${formatNumber(currentTotal)} contratos vigentes monitorados.`;

    const calloutParts = [
      `${formatNumber(summary.comResponsaveisCompletos)} com gestor e fiscal identificados`,
      `${formatNumber(withoutFiscal)} sem fiscal designado`,
    ];

    if (topOrganization?.organization) {
      calloutParts.push(`${topOrganization.organization} concentra ${formatNumber(topOrganization.count)} registros`);
    }

    elements.heroCallout.textContent = calloutParts.join(" · ");
    elements.heroCallout.classList.toggle("hidden", calloutParts.length === 0);
  }

  function renderStatusGrid() {
    const summary = state.payload.summary;
    const cards = [
      {
        label: "Em validação",
        value: formatNumber(summary.emAcompanhamento),
        detail: "Vigência sem confirmação final",
        className: "status-card--warning",
      },
      {
        label: "Sem gestor designado",
        value: formatNumber(summary.semGestor),
        detail: "Providência necessária",
        className: "status-card--danger",
      },
      {
        label: "Sem fiscal designado",
        value: formatNumber(summary.semFiscal),
        detail: "Providência necessária",
        className: "status-card--danger",
      },
      {
        label: "Designações completas",
        value: formatNumber(summary.comResponsaveisCompletos),
        detail: "Gestor e fiscal identificados",
        className: "status-card--success",
      },
    ];

    setHtml(
      elements.statusGrid,
      cards
      .map(
        (card) => `
          <article class="status-card ${card.className}">
            <span class="eyebrow">${escapeHtml(card.label)}</span>
            <strong class="status-value">${escapeHtml(card.value)}</strong>
            <span class="status-detail">${escapeHtml(card.detail)}</span>
          </article>
        `
      )
      .join("")
    );
  }

  function renderCoverageGrid() {
    const summary = state.payload.summary;
    const currentTotal = summary.contratosAtuais || 0;
    renderCardCollection(elements.coverageGrid, [
      {
        label: "Cobertura completa",
        value: formatPercent(summary.comResponsaveisCompletos, currentTotal),
        detail: `${formatNumber(summary.comResponsaveisCompletos)} com gestor e fiscal`,
      },
      {
        label: "Sem gestor",
        value: formatPercent(summary.semGestor, currentTotal),
        detail: `${formatNumber(summary.semGestor)} registros`,
      },
      {
        label: "Sem fiscal",
        value: formatPercent(summary.semFiscal, currentTotal),
        detail: `${formatNumber(summary.semFiscal)} registros`,
      },
      {
        label: "Sem responsáveis",
        value: formatPercent(summary.semGestorEFiscal, currentTotal),
        detail: `${formatNumber(summary.semGestorEFiscal)} registros`,
      },
    ]);
  }

  function renderSourceGrid() {
    const summary = state.payload.summary;
    const currentTotal = summary.contratosAtuais || 0;
    renderCardCollection(elements.sourceGrid, [
      {
        label: "Cruzamento confirmado",
        value: formatPercent(summary.cruzados, currentTotal),
        detail: `${formatNumber(summary.cruzados)} registros`,
      },
      {
        label: "Apenas Diário Oficial",
        value: formatPercent(summary.somenteDiario, currentTotal),
        detail: `${formatNumber(summary.somenteDiario)} registros`,
      },
      {
        label: "Apenas Portal",
        value: formatPercent(summary.somentePortal, currentTotal),
        detail: `${formatNumber(summary.somentePortal)} registros`,
      },
      {
        label: "Registros analisados",
        value: formatNumber(summary.analisados),
        detail: "Eventos considerados na leitura",
      },
    ]);
  }

  function renderDeadlineGrid() {
    const currentRecords = getCurrentRecords();
    const withDeadline = currentRecords.filter((record) => record._endDate);
    const expiring30 = withDeadline.filter((record) => record.vigency?.daysUntilEnd != null && record.vigency.daysUntilEnd >= 0 && record.vigency.daysUntilEnd <= 30);
    const expiring90 = withDeadline.filter((record) => record.vigency?.daysUntilEnd != null && record.vigency.daysUntilEnd > 30 && record.vigency.daysUntilEnd <= 90);
    const overdue = withDeadline.filter((record) => record.vigency?.daysUntilEnd != null && record.vigency.daysUntilEnd < 0);

    renderCardCollection(elements.deadlineGrid, [
      {
        label: "Prazo identificado",
        value: formatNumber(withDeadline.length),
        detail: formatPercent(withDeadline.length, currentRecords.length),
      },
      {
        label: "Até 30 dias",
        value: formatNumber(expiring30.length),
        detail: "Encerramento mais próximo",
      },
      {
        label: "31 a 90 dias",
        value: formatNumber(expiring90.length),
        detail: "Monitoramento intermediário",
      },
      {
        label: "Prazo expirado",
        value: formatNumber(overdue.length),
        detail: "Exige conferência documental",
      },
    ]);
  }

  function getPendingResponsibilityWithDeadlineRecords(records = getCurrentRecords()) {
    return records
      .filter(
        (record) =>
          (record.managementState === "sem_gestor" ||
            record.managementState === "sem_fiscal" ||
            record.managementState === "sem_gestor_e_fiscal") &&
          record._endDate
      )
      .sort((a, b) => {
        const aDays = a.vigency?.daysUntilEnd ?? Number.MAX_SAFE_INTEGER;
        const bDays = b.vigency?.daysUntilEnd ?? Number.MAX_SAFE_INTEGER;
        if (aDays !== bDays) return aDays - bDays;
        return compareByPriority(a, b);
      });
  }

  function renderPendingDeadlineAnalysis() {
    const currentRecords = getCurrentRecords();
    const pendingWithDeadline = getPendingResponsibilityWithDeadlineRecords(currentRecords);
    const withoutManager = pendingWithDeadline.filter((record) => record.managementState === "sem_gestor");
    const withoutInspector = pendingWithDeadline.filter((record) => record.managementState === "sem_fiscal");
    const withoutBoth = pendingWithDeadline.filter((record) => record.managementState === "sem_gestor_e_fiscal");

    elements.pendingDeadlineSummary.textContent =
      pendingWithDeadline.length > 0
        ? `${formatNumber(pendingWithDeadline.length)} contrato(s) sem gestor ou fiscal têm prazo final identificado.`
        : "Nenhum contrato sem gestor ou fiscal tem prazo final identificado.";

    if (!pendingWithDeadline.length) {
      setHtml(elements.pendingDeadlineRecords, `<div class="empty-state">Sem pendências com prazo identificado.</div>`);
      return;
    }

    setHtml(
      elements.pendingDeadlineRecords,
      `
        <article class="compact-card">
          <strong>${formatNumber(withoutManager.length)}</strong>
          <span>Sem gestor com prazo</span>
        </article>
        <article class="compact-card">
          <strong>${formatNumber(withoutInspector.length)}</strong>
          <span>Sem fiscal com prazo</span>
        </article>
        <article class="compact-card">
          <strong>${formatNumber(withoutBoth.length)}</strong>
          <span>Sem responsáveis com prazo</span>
        </article>
        ${pendingWithDeadline
          .slice(0, 6)
          .map(
            (record) => `
              <article class="compact-record">
                <div class="compact-record-copy">
                  <strong>${escapeHtml(record.contractNumber || "Contrato sem número")}</strong>
                  <span>${escapeHtml(getLabel("management", record.managementState))}</span>
                  <small>${escapeHtml(record.organization || "Órgão não informado")}</small>
                </div>
                <div class="compact-record-meta">
                  <strong>${escapeHtml(formatDate(record._endDate))}</strong>
                  <small>${escapeHtml(formatDayDelta(record.vigency?.daysUntilEnd))}</small>
                </div>
              </article>
            `
          )
          .join("")}
      `
    );
  }

  function renderOrganizationSummary() {
    const list = state.payload.organizationSummary || [];
    if (!list.length) {
      setHtml(elements.organizationSummary, `<div class="empty-state">Nenhum registro disponível.</div>`);
      return;
    }

    const maxCount = Math.max(...list.map((item) => item.count), 1);
    const currentTotal = state.payload.summary?.contratosAtuais || 1;

    setHtml(
      elements.organizationSummary,
      list
      .map(
        (item) => `
          <div class="organization-row">
            <div class="organization-head">
              <div>
                <strong>${escapeHtml(item.organization)}</strong>
                <small>${escapeHtml(formatPercent(item.count, currentTotal))} do total vigente</small>
              </div>
              <span>${escapeHtml(formatNumber(item.count))}</span>
            </div>
            <div class="organization-bar">
              <span style="width: ${Math.max(10, (item.count / maxCount) * 100)}%"></span>
            </div>
          </div>
        `
      )
      .join("")
    );
  }

  function sampleContracts(records) {
    return records
      .slice(0, 3)
      .map((record) => record.contractNumber || record.normalizedKey || "Sem número")
      .join(" · ");
  }

  function renderPriorityGroups() {
    const currentRecords = getCurrentRecords();
    const groups = [
      {
        preset: "semGestorEFiscal",
        title: "Sem responsáveis",
        count: currentRecords.filter((record) => record.managementState === "sem_gestor_e_fiscal").length,
        sample: sampleContracts(currentRecords.filter((record) => record.managementState === "sem_gestor_e_fiscal")),
      },
      {
        preset: "semFiscal",
        title: "Sem fiscal",
        count: currentRecords.filter((record) => record.managementState === "sem_fiscal").length,
        sample: sampleContracts(currentRecords.filter((record) => record.managementState === "sem_fiscal")),
      },
      {
        preset: "altaCriticidade",
        title: "Alta criticidade",
        count: currentRecords.filter((record) => record._criticality === "alta").length,
        sample: sampleContracts(currentRecords.filter((record) => record._criticality === "alta")),
      },
      {
        preset: "somenteDiario",
        title: "Apenas Diário Oficial",
        count: currentRecords.filter((record) => record.sourceStatus === "somente_diario").length,
        sample: sampleContracts(currentRecords.filter((record) => record.sourceStatus === "somente_diario")),
      },
    ];

    setHtml(
      elements.priorityGroups,
      groups
      .map(
        (group) => `
          <article class="priority-card">
            <span class="eyebrow">Consulta prioritária</span>
            <strong>${escapeHtml(group.title)}</strong>
            <span class="priority-count">${escapeHtml(formatNumber(group.count))}</span>
            <div class="priority-sample">${escapeHtml(group.sample || "Nenhum registro disponível.")}</div>
            <button type="button" data-preset="${escapeHtml(group.preset)}">Abrir</button>
          </article>
        `
      )
      .join("")
    );
  }

  function renderInsights() {
    const currentRecords = getCurrentRecords();
    const summary = summarizeRecordSet(currentRecords);
    const topOrganization = getTopOrganization(currentRecords);
    const latestRecord = [...currentRecords].sort((a, b) => b._movementTimestamp - a._movementTimestamp)[0];

    const insights = [
      {
        title: "Cobertura de responsáveis",
        text: `${formatPercent(summary.withCompleteAssignments, summary.total)} dos contratos vigentes têm gestor e fiscal identificados.`,
      },
      {
        title: "Fiscalização",
        text: `${formatNumber(summary.withoutInspector)} contratos seguem sem fiscal identificado.`,
      },
      {
        title: "Base documental",
        text: `${formatPercent(summary.diaryOnly, summary.total)} dos registros vigentes dependem apenas do Diário Oficial.`,
      },
      {
        title: "Concentração institucional",
        text: topOrganization
          ? `${topOrganization.organization} concentra ${formatNumber(topOrganization.count)} contratos vigentes.`
          : "Sem concentração identificada.",
      },
      {
        title: "Prazos próximos",
        text:
          summary.expiringSoon > 0
            ? `${formatNumber(summary.expiringSoon)} contratos têm prazo final em até 30 dias.`
            : "Nenhum contrato com prazo conhecido vence em até 30 dias.",
      },
      {
        title: "Última movimentação",
        text: latestRecord
          ? `${latestRecord.contractNumber || "Registro sem número"} teve atualização em ${formatDate(latestRecord._movementDate)}.`
          : "Sem movimentações recentes identificadas.",
      },
    ];

    setHtml(
      elements.insightList,
      insights
      .map(
        (insight) => `
          <article class="insight-item">
            <strong>${escapeHtml(insight.title)}</strong>
            <p>${escapeHtml(insight.text)}</p>
          </article>
        `
      )
      .join("")
    );
  }

  function renderRecentMovements() {
    const records = [...getCurrentRecords()]
      .filter((record) => record._movementTimestamp > 0)
      .sort((a, b) => b._movementTimestamp - a._movementTimestamp)
      .slice(0, 6);

    if (!records.length) {
      setHtml(elements.recentMovements, `<div class="empty-state">Nenhuma movimentação disponível.</div>`);
      return;
    }

    setHtml(
      elements.recentMovements,
      records
      .map(
        (record) => `
          <article class="movement-item">
            <div class="movement-copy">
              <strong>${escapeHtml(record.contractNumber || "Contrato sem número")}</strong>
              <span>${escapeHtml(record.organization || "Órgão não informado")}</span>
              <small>${escapeHtml(record.lastMovementTitle || "Movimentação registrada")}</small>
            </div>
            <time datetime="${escapeHtml(record.managementActAt || record.publishedAt || "")}">${escapeHtml(formatDate(record._movementDate))}</time>
          </article>
        `
      )
      .join("")
    );
  }

  function renderQuickPresets() {
    setHtml(
      elements.quickPresets,
      Object.entries(PRESETS)
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
        .join("")
    );
  }

  function getPersonDisplay(person) {
    if (person?.name) {
      return {
        title: person.name,
        subtitle: person.role || "Responsável designado",
      };
    }

    if (person?.needsReview) {
      return {
        title: "Em validação",
        subtitle: "Identificação pendente",
      };
    }

    return {
      title: "Não informado",
      subtitle: "Sem registro",
    };
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

  function getBadgeToneByCriticality(value) {
    if (value === "alta") return "danger";
    if (value === "media") return "warning";
    return "primary";
  }

  function getToneClass(record) {
    if (record._criticality === "alta") return "record-card--critical";
    if (record._criticality === "media") return "record-card--warning";
    return "";
  }

  function getAssignmentTone(person, personnelStatus, exonerationSignal) {
    if (exonerationSignal || personnelStatus === "exonerado") return "danger";
    if (person?.needsReview) return "warning";
    if (person?.name && personnelStatus === "ativo") return "success";
    if (person?.name) return "primary";
    return "muted";
  }

  function getAssignmentStatusLabel(person, personnelStatus, exonerationSignal) {
    if (exonerationSignal || personnelStatus === "exonerado") return "Exoneracao identificada";
    if (person?.needsReview) return "Em revisao";
    if (person?.name && personnelStatus === "ativo") return "Servidor ativo";
    if (person?.name) return "Designacao localizada";
    return "Sem designacao";
  }

  function getAdditiveDisplay(record) {
    const additiveCount = Number(record.additives?.totalKnown || 0);
    const apostilleCount = Number(record.additives?.apostilleCount || 0);
    const terminationCount = Number(record.additives?.terminationCount || 0);
    const lifecycleSummary = record.lifecycle?.summary || "Sem movimentacao complementar localizada";

    if (record.additives?.hasActiveTermination || terminationCount > 0) {
      return {
        title: "Rescisao localizada",
        detail: lifecycleSummary,
      };
    }

    if (additiveCount > 0) {
      return {
        title: `${formatNumber(additiveCount)} evento(s)`,
        detail: lifecycleSummary,
      };
    }

    if (apostilleCount > 0) {
      return {
        title: `${formatNumber(apostilleCount)} apostila(s)`,
        detail: lifecycleSummary,
      };
    }

    return {
      title: "Sem aditivos",
      detail: lifecycleSummary,
    };
  }

  function getReviewDisplay(record) {
    if (!record.review?.required) {
      return {
        title: "Sem revisao aberta",
        detail: "Registro consolidado pela automacao atual",
        tone: "success",
      };
    }

    const detailParts = [];
    if (record.review?.reasonCount > 0) detailParts.push(`${formatNumber(record.review.reasonCount)} motivo(s)`);
    if (record.review?.candidateCount > 0) detailParts.push(`${formatNumber(record.review.candidateCount)} cruzamento(s)`);
    if (record.review?.divergenceCount > 0) detailParts.push(`${formatNumber(record.review.divergenceCount)} divergencia(s)`);

    return {
      title: record.review?.reasonSummary || "Revisao dirigida",
      detail: detailParts.join(" · ") || "Conferencia focal necessaria",
      tone: record.review?.priority === "alta" ? "danger" : "warning",
    };
  }

  function renderAssignmentCard(label, person, personnelStatus, exonerationSignal) {
    const display = getPersonDisplay(person);
    const tone = getAssignmentTone(person, personnelStatus, exonerationSignal);
    const statusLabel = getAssignmentStatusLabel(person, personnelStatus, exonerationSignal);

    let detail = display.subtitle;
    if (person?.assignedAt) {
      detail = `Ato em ${formatDate(person.assignedAt)}`;
    }
    if (exonerationSignal || personnelStatus === "exonerado") {
      detail = "Substituicao atual nao confirmada";
    }

    return `
      <article class="assignment-card assignment-card--${escapeHtml(tone)}">
        <div class="assignment-card-head">
          <span class="assignment-label">${escapeHtml(label)}</span>
          <span class="assignment-status assignment-status--${escapeHtml(tone)}">${escapeHtml(statusLabel)}</span>
        </div>
        <strong>${escapeHtml(display.title)}</strong>
        <small>${escapeHtml(detail)}</small>
      </article>
    `;
  }

  function renderEvidenceCard(label, title, detail, tone = "primary") {
    return `
      <article class="evidence-card evidence-card--${escapeHtml(tone)}">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(title)}</strong>
        <small>${escapeHtml(detail)}</small>
      </article>
    `;
  }

  function renderBadges(record) {
    const badges = [
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
      {
        tone: getBadgeToneByCriticality(record._criticality),
        label: getLabel("criticality", record._criticality),
      },
    ];

    if (record.additives?.isAdditivado || Number(record.additives?.totalKnown || 0) > 0) {
      badges.push({ tone: "primary", label: "Com aditivos" });
    }

    if (record.review?.required) {
      badges.push({
        tone: record.review?.priority === "alta" ? "danger" : "warning",
        label: "Revisao dirigida",
      });
    }

    return badges
      .map((badge) => `<span class="badge badge--${escapeHtml(badge.tone)}">${escapeHtml(badge.label)}</span>`)
      .join("");
  }

  function renderAlertPills(record) {
    const alerts = getUniqueAlerts(record.alerts).slice(0, 5);
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
      ? `<a href="${escapeHtml(record.links.diary)}" target="_blank" rel="noopener noreferrer">Diário Oficial</a>`
      : "";
    const portalLink = record.links?.portal
      ? `<a href="${escapeHtml(record.links.portal)}" target="_blank" rel="noopener noreferrer">Portal da Transparência</a>`
      : "";

    return `
      <article class="record-card ${getToneClass(record)}">
        <div class="record-head">
          <div class="record-heading">
            <span class="record-number">${escapeHtml(record.contractNumber || "Contrato sem número")}</span>
            <h3>${escapeHtml(record.organization || "Órgão não identificado")}</h3>
            <span class="record-summary">${escapeHtml(record.managementSummary || "Síntese não disponível")}</span>
          </div>
          <div class="badge-row">
            ${renderBadges(record)}
          </div>
        </div>

        <p class="record-object">${escapeHtml(truncateText(record.object || "Objeto não informado"))}</p>

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
            <span>Situação contratual</span>
            <strong>${escapeHtml(record.vigency?.label || "Vigência não informada")}</strong>
            <small>${escapeHtml(record.vigency?.sourceLabel || "Sem detalhamento")}</small>
          </div>
          <div class="meta-block">
            <span>Prazo final</span>
            <strong>${escapeHtml(record._endDate ? formatDate(record._endDate) : "Data não informada")}</strong>
            <small>${escapeHtml(formatDayDelta(record.vigency?.daysUntilEnd))}</small>
          </div>
        </div>

        <div class="record-meta">
          <div class="meta-block">
            <span>Fornecedor</span>
            <strong>${escapeHtml(record.supplier || "Não informado")}</strong>
            <small>${escapeHtml(record.valueLabel || formatCurrency(record._valueNumber))}</small>
          </div>
          <div class="meta-block">
            <span>Gestão</span>
            <strong>${escapeHtml(record.administration || "Não informado")}</strong>
            <small>${escapeHtml(record.year ? `Ano ${record.year}` : "Ano não informado")}</small>
          </div>
          <div class="meta-block">
            <span>Base documental</span>
            <strong>${escapeHtml(getLabel("source", record.sourceStatus))}</strong>
            <small>${escapeHtml(`${formatNumber(record.movementCount || 0)} movimentação(ões)`)}</small>
          </div>
          <div class="meta-block">
            <span>Última movimentação</span>
            <strong>${escapeHtml(formatDate(record._movementDate))}</strong>
            <small>${escapeHtml(record.lastMovementTitle || "Sem detalhamento")}</small>
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

  function renderRecordCardV2(record) {
    const additive = getAdditiveDisplay(record);
    const review = getReviewDisplay(record);
    const contractMeta = [
      record.processNumber ? `Processo ${record.processNumber}` : "",
      record.administration || "",
      record.year ? `Ano ${record.year}` : "",
    ]
      .filter(Boolean)
      .join(" · ");
    const sourceDetailParts = [];
    if (record.hasDiary) sourceDetailParts.push(`${formatNumber(record.movementCount || 0)} ato(s) no Diario`);
    if (record.hasOfficialPortal) sourceDetailParts.push("Cadastro localizado no portal");
    if (record.review?.candidateCount > 0) {
      sourceDetailParts.push(`${formatNumber(record.review.candidateCount)} cruzamento(s) pendente(s)`);
    }
    const objectText = record.object || record.lastMovementTitle || "Objeto nao localizado";
    const diaryLink = record.links?.diary
      ? `<a href="${escapeHtml(record.links.diary)}" target="_blank" rel="noopener noreferrer">Diario Oficial</a>`
      : "";
    const portalLink = record.links?.portal
      ? `<a href="${escapeHtml(record.links.portal)}" target="_blank" rel="noopener noreferrer">Portal da Transparencia</a>`
      : "";

    return `
      <article class="record-card ${getToneClass(record)}">
        <div class="record-head">
          <div class="record-heading">
            <span class="record-number">${escapeHtml(record.contractNumber || "Contrato sem numero")}</span>
            <h3>${escapeHtml(record.organization || "Orgao nao identificado")}</h3>
            ${contractMeta ? `<div class="record-kicker">${escapeHtml(contractMeta)}</div>` : ""}
            <span class="record-summary">${escapeHtml(record.managementSummary || "Sintese nao disponivel")}</span>
          </div>
          <div class="badge-row">
            ${renderBadges(record)}
          </div>
        </div>

        <div class="record-overview">
          <article class="overview-block overview-block--object">
            <span class="overview-label">Objeto / servico</span>
            <p class="record-object">${escapeHtml(truncateText(objectText, 420))}</p>
          </article>
          <article class="overview-block">
            <span class="overview-label">Fornecedor</span>
            <strong>${escapeHtml(record.supplier || "Nao localizado")}</strong>
            <small>${escapeHtml(record.valueLabel || formatCurrency(record._valueNumber))}</small>
          </article>
          <article class="overview-block">
            <span class="overview-label">Prazo atual</span>
            <strong>${escapeHtml(record._endDate ? formatDate(record._endDate) : "Prazo nao localizado")}</strong>
            <small>${escapeHtml(formatDayDelta(record.vigency?.daysUntilEnd))}</small>
          </article>
          <article class="overview-block">
            <span class="overview-label">Aditivos</span>
            <strong>${escapeHtml(additive.title)}</strong>
            <small>${escapeHtml(additive.detail)}</small>
          </article>
        </div>

        <div class="record-section-grid">
          <section class="record-section">
            <div class="section-mini-head">
              <span class="eyebrow">Responsaveis</span>
              <strong>Gestao contratual</strong>
            </div>
            <div class="assignment-grid">
              ${renderAssignmentCard("Gestor", record.manager, record.managerPersonnelStatus, record.managerExonerationSignal)}
              ${renderAssignmentCard("Fiscal", record.inspector, record.inspectorPersonnelStatus, record.inspectorExonerationSignal)}
            </div>
          </section>

          <section class="record-section">
            <div class="section-mini-head">
              <span class="eyebrow">Evidencias</span>
              <strong>Rastreio do contrato</strong>
            </div>
            <div class="evidence-grid">
              ${renderEvidenceCard(
                "Situacao contratual",
                record.vigency?.label || "Vigencia nao informada",
                record.vigency?.sourceLabel || "Sem detalhamento",
                getBadgeToneByVigency(record.vigency?.state)
              )}
              ${renderEvidenceCard(
                "Base documental",
                getLabel("source", record.sourceStatus),
                sourceDetailParts.join(" · ") || "Sem rastreio complementar",
                getBadgeToneBySource(record.sourceStatus)
              )}
              ${renderEvidenceCard(
                "Ultima movimentacao",
                formatDate(record._movementDate),
                record.lastMovementTitle || "Sem detalhamento",
                "primary"
              )}
              ${renderEvidenceCard("Revisao", review.title, review.detail, review.tone)}
            </div>
          </section>
        </div>

        ${renderAlertPills(record)}

        <div class="record-actions">
          ${diaryLink}
          ${portalLink}
        </div>
      </article>
    `;
  }

  function renderAlertSummary() {
    const currentRecords = getCurrentRecords();
    const alertRecords = currentRecords.filter((record) => getUniqueAlerts(record.alerts).length > 0);
    const summary = summarizeRecordSet(alertRecords);

    elements.alertsSummary.textContent = `${formatNumber(alertRecords.length)} registros com ocorrência e ordenação por prioridade operacional.`;

    renderCardCollection(elements.alertSummaryGrid, [
      {
        label: "Alta criticidade",
        value: formatNumber(summary.critical),
        detail: "Maior peso operacional",
      },
      {
        label: "Sem responsáveis",
        value: formatNumber(summary.withoutBoth),
        detail: "Gestor e fiscal ausentes",
      },
      {
        label: "Sem fiscal",
        value: formatNumber(summary.withoutInspector),
        detail: "Pendência de fiscalização",
      },
      {
        label: "Pendências com prazo",
        value: formatNumber(getPendingResponsibilityWithDeadlineRecords(alertRecords).length),
        detail: "Sem gestor ou fiscal com vigência identificada",
      },
      {
        label: "Apenas Diário Oficial",
        value: formatNumber(summary.diaryOnly),
        detail: "Sem confirmação cruzada",
      },
    ]);
  }

  function renderAlertRecords() {
    const records = [...getCurrentRecords()]
      .filter((record) => getUniqueAlerts(record.alerts).length > 0)
      .sort(compareByPriority)
      .slice(0, 12);

    if (!records.length) {
      setHtml(elements.alertRecords, `<div class="empty-state">Nenhuma ocorrência prioritária.</div>`);
      return;
    }

    setHtml(elements.alertRecords, records.map(renderRecordCard).join(""));
  }

  function renderReviewQueue() {
    const summary = state.payload.reviewSummary || {};
    const queue = state.payload.reviewQueue || [];

    elements.reviewQueueSummary.textContent = queue.length
      ? `${formatNumber(summary.current || 0)} contrato(s) atuais em revisao dirigida.`
      : "Sem revisao dirigida aberta.";

    renderCardCollection(elements.reviewSummaryGrid, [
      {
        label: "Fila atual",
        value: formatNumber(summary.current),
        detail: `${formatNumber(summary.total)} itens no total`,
      },
      {
        label: "Prioridade alta",
        value: formatNumber(summary.high),
        detail: "Casos mais sensiveis",
      },
      {
        label: "Cruzamento pendente",
        value: formatNumber(summary.crossPending),
        detail: "Correspondencia manual",
      },
      {
        label: "Divergencias",
        value: formatNumber(summary.divergent),
        detail: "Fontes em conflito",
      },
    ]);

    if (!queue.length) {
      setHtml(elements.reviewQueue, `<div class="empty-state">Sem fila de revisao.</div>`);
      return;
    }

    setHtml(
      elements.reviewQueue,
      queue
        .slice(0, 8)
        .map((item) => {
          const stateLabel = item.isCurrent ? "Atual" : "Historico";
          const secondary = [];
          if (item.candidateCount > 0) secondary.push(`${formatNumber(item.candidateCount)} candidato(s)`);
          if (item.divergenceCount > 0) secondary.push(`${formatNumber(item.divergenceCount)} divergencia(s)`);
          if (item.daysUntilEnd != null) secondary.push(formatDayDelta(item.daysUntilEnd));

          return `
            <article class="compact-record">
              <div class="compact-record-copy">
                <strong>${escapeHtml(item.contractNumber || item.normalizedKey || "Contrato sem numero")}</strong>
                <span>${escapeHtml(item.reasonSummary || "Revisao dirigida")}</span>
                <small>${escapeHtml(item.organization || "Orgao nao informado")}</small>
                <small>${escapeHtml(secondary.join(" · ") || "Sem complemento")}</small>
              </div>
              <div class="compact-record-meta">
                <strong>${escapeHtml(item.priority || "baixa")}</strong>
                <small>${escapeHtml(stateLabel)}</small>
                <small>${escapeHtml(item.sourceAlignment || "parcial")}</small>
              </div>
            </article>
          `;
        })
        .join("")
    );
  }

  function getActiveFiltersText() {
    const activeFilters = [];
    if (state.filters.query) activeFilters.push(`Busca: ${state.filters.query}`);
    if (state.filters.organization) activeFilters.push(`Órgão: ${state.filters.organization}`);
    if (state.filters.administration) activeFilters.push(`Gestão: ${state.filters.administration}`);
    if (state.filters.vigency !== "todos") activeFilters.push(`Vigência: ${getLabel("vigency", state.filters.vigency)}`);
    if (state.filters.management !== "todos") activeFilters.push(`Responsáveis: ${getLabel("management", state.filters.management)}`);
    if (state.filters.source !== "todos") activeFilters.push(`Fonte: ${getLabel("source", state.filters.source)}`);
    if (state.filters.criticality !== "todos") activeFilters.push(`Criticidade: ${getLabel("criticality", state.filters.criticality)}`);
    if (state.filters.scope !== DEFAULT_FILTERS.scope) activeFilters.push(`Escopo: ${getLabel("scope", state.filters.scope)}`);
    if (state.filters.sort !== DEFAULT_FILTERS.sort) activeFilters.push(`Ordenação: ${getLabel("sort", state.filters.sort)}`);
    return activeFilters;
  }

  function renderResultInsights(records) {
    const summary = summarizeRecordSet(records);
    const withAdditives = records.filter(
      (record) => record.additives?.isAdditivado || Number(record.additives?.totalKnown || 0) > 0
    ).length;
    const topOrganization = getTopOrganization(records);

    renderCardCollection(elements.resultInsightGrid, [
      {
        label: "Registros no recorte",
        value: formatNumber(summary.total),
        detail: "Resultado atual da consulta",
      },
      {
        label: "Designações completas",
        value: formatNumber(summary.withCompleteAssignments),
        detail: formatPercent(summary.withCompleteAssignments, summary.total),
      },
      {
        label: "Alta criticidade",
        value: formatNumber(summary.critical),
        detail: formatPercent(summary.critical, summary.total),
      },
      {
        label: "Maior concentração",
        value: formatNumber(topOrganization?.count || 0),
        detail: topOrganization?.organization || "Sem concentração",
      },
    ]);
  }

  function renderResultInsightsV2(records) {
    const summary = summarizeRecordSet(records);
    const withAdditives = records.filter(
      (record) => record.additives?.isAdditivado || Number(record.additives?.totalKnown || 0) > 0
    ).length;

    renderCardCollection(elements.resultInsightGrid, [
      {
        label: "Registros no recorte",
        value: formatNumber(summary.total),
        detail: "Resultado atual da consulta",
      },
      {
        label: "Prazo identificado",
        value: formatNumber(summary.withDeadline),
        detail: formatPercent(summary.withDeadline, summary.total),
      },
      {
        label: "Com aditivos",
        value: formatNumber(withAdditives),
        detail: formatPercent(withAdditives, summary.total),
      },
      {
        label: "Designacoes completas",
        value: formatNumber(summary.withCompleteAssignments),
        detail: formatPercent(summary.withCompleteAssignments, summary.total),
      },
    ]);
  }

  function renderResults() {
    const filtered = getFilteredRecords();
    const visible = filtered.slice(0, state.visibleCount);

    elements.resultsMeta.textContent = `${formatNumber(filtered.length)} registros encontrados · ${getLabel("sort", state.filters.sort)}`;

    elements.resultsMeta.textContent = `${formatNumber(filtered.length)} ficha(s) contratuais · ${getLabel("sort", state.filters.sort)}`;

    const activeFilters = getActiveFiltersText();
    elements.activeFilterSummary.textContent = activeFilters.join(" · ");
    elements.activeFilterSummary.classList.toggle("hidden", activeFilters.length === 0);

    renderResultInsightsV2(filtered);

    if (!visible.length) {
      setHtml(elements.recordList, `<div class="empty-state">Nenhum registro encontrado.</div>`);
      elements.loadMore.classList.add("hidden");
      return;
    }

    setHtml(elements.recordList, visible.map(renderRecordCardV2).join(""));
    elements.loadMore.classList.toggle("hidden", visible.length >= filtered.length);
  }

  function renderFooter() {
    const summary = state.payload.summary || {};
    elements.footerCopy.textContent = `${formatNumber(summary.contratosAtuais || 0)} contratos vigentes monitorados · atualização ${formatDateTime(state.payload.generatedAt)}`;
  }

  function renderAll() {
    if (!state.payload) return;
    renderHero();
    renderSummaryCardsV2();
    renderMethodology();
    renderStatusGrid();
    renderCoverageGrid();
    renderSourceGrid();
    renderDeadlineGrid();
    renderPendingDeadlineAnalysis();
    renderOrganizationSummary();
    renderRecentMovements();
    renderInsights();
    renderPriorityGroups();
    renderQuickPresets();
    renderAlertSummary();
    renderAlertRecords();
    renderReviewQueue();
    renderResults();
    renderFooter();
    syncUrlState();
  }

  function resetFilters() {
    state.filters = { ...DEFAULT_FILTERS };
    state.visibleCount = PAGE_SIZE;
    syncControls();
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

  function updateFilter(key, value) {
    if (key === "query") {
      state.filters[key] = sanitizeQueryInput(value);
    } else if (key === "organization" || key === "administration") {
      state.filters[key] = sanitizePlainText(value, 160);
    } else {
      state.filters[key] = value;
    }
    state.visibleCount = PAGE_SIZE;
    renderAll();
  }

  function scheduleQueryUpdate(value) {
    window.clearTimeout(state.searchTimer);
    state.searchTimer = window.setTimeout(() => {
      updateFilter("query", value.trim());
    }, SEARCH_DEBOUNCE_MS);
  }

  function csvEscape(value) {
    return `"${String(value ?? "").replace(/"/g, '""')}"`;
  }

  function buildCsv(records) {
    const headers = [
      "Contrato",
      "Órgão",
      "Gestão",
      "Ano",
      "Fornecedor",
      "Objeto",
      "Gestor",
      "Cargo do gestor",
      "Fiscal",
      "Cargo do fiscal",
      "Situação contratual",
      "Situação dos responsáveis",
      "Fonte",
      "Criticidade",
      "Prazo final",
      "Dias para o término",
      "Última movimentação",
      "Título da movimentação",
      "Valor",
      "Movimentações",
      "Link do Diário Oficial",
      "Link do Portal da Transparência",
      "Alertas",
    ];

    const lines = records.map((record) => [
      record.contractNumber || "",
      record.organization || "",
      record.administration || "",
      record.year || "",
      record.supplier || "",
      truncateText(record.object || "", 400),
      record.manager?.name || "",
      record.manager?.role || "",
      record.inspector?.name || "",
      record.inspector?.role || "",
      getLabel("vigency", record.vigency?.state),
      getLabel("management", record.managementState),
      getLabel("source", record.sourceStatus),
      getLabel("criticality", record._criticality),
      record._endDate ? formatDate(record._endDate) : "",
      record.vigency?.daysUntilEnd ?? "",
      formatDate(record._movementDate),
      record.lastMovementTitle || "",
      record.valueLabel || formatCurrency(record._valueNumber),
      record.movementCount || 0,
      record.links?.diary || "",
      record.links?.portal || "",
      getUniqueAlerts(record.alerts)
        .map((alert) => alert.title)
        .join(" | "),
    ]);

    return [headers, ...lines].map((columns) => columns.map(csvEscape).join(";")).join("\r\n");
  }

  function downloadFile(filename, content, type) {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function flashButton(button, label) {
    const original = button.dataset.originalLabel || button.textContent;
    button.dataset.originalLabel = original;
    button.textContent = label;
    window.setTimeout(() => {
      button.textContent = original;
    }, 1600);
  }

  async function copyToClipboard(text) {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }

    const input = document.createElement("textarea");
    input.value = text;
    input.setAttribute("readonly", "readonly");
    input.style.position = "fixed";
    input.style.opacity = "0";
    document.body.appendChild(input);
    input.select();
    const success = document.execCommand("copy");
    input.remove();
    return success;
  }

  async function handleShareView() {
    syncUrlState();
    try {
      await copyToClipboard(window.location.href);
      flashButton(elements.shareView, "Link copiado");
    } catch {
      flashButton(elements.shareView, "Não foi possível copiar");
    }
  }

  function handleExportCsv() {
    const records = getFilteredRecords();
    if (!records.length) {
      flashButton(elements.exportCsv, "Sem dados");
      return;
    }

    const generatedAt = state.payload.generatedAt ? new Date(state.payload.generatedAt) : new Date();
    const stamp = `${generatedAt.getFullYear()}-${String(generatedAt.getMonth() + 1).padStart(2, "0")}-${String(generatedAt.getDate()).padStart(2, "0")}`;
    downloadFile(`contratos-vigentes-${stamp}.csv`, buildCsv(records), "text/csv;charset=utf-8");
    flashButton(elements.exportCsv, "CSV gerado");
  }

  function bindEvents() {
    elements.viewButtons.forEach((button) => {
      button.addEventListener("click", () => {
        setView(button.dataset.viewButton);
      });
    });

    elements.searchInput.addEventListener("input", () => {
      scheduleQueryUpdate(elements.searchInput.value);
    });

    [
      [elements.scopeSelect, "scope"],
      [elements.organizationSelect, "organization"],
      [elements.administrationSelect, "administration"],
      [elements.vigencySelect, "vigency"],
      [elements.managementSelect, "management"],
      [elements.sourceSelect, "source"],
      [elements.criticalitySelect, "criticality"],
      [elements.sortSelect, "sort"],
    ].forEach(([element, key]) => {
      element.addEventListener("change", () => {
        updateFilter(key, element.value);
      });
    });

    elements.clearFilters.addEventListener("click", resetFilters);
    elements.shareView.addEventListener("click", handleShareView);
    elements.exportCsv.addEventListener("click", handleExportCsv);

    elements.loadMore.addEventListener("click", () => {
      state.visibleCount += PAGE_SIZE;
      renderResults();
    });

    document.addEventListener("click", (event) => {
      const presetButton = event.target.closest("[data-preset]");
      if (!presetButton) return;
      applyPreset(presetButton.dataset.preset);
    });

    document.addEventListener("click", (event) => {
      const viewButton = event.target.closest("[data-open-view]");
      if (!viewButton) return;
      setView(viewButton.dataset.openView);
    });

    window.addEventListener("popstate", () => {
      const nextState = readUrlState();
      state.view = nextState.view;
      state.filters = { ...DEFAULT_FILTERS, ...nextState.filters };
      state.visibleCount = PAGE_SIZE;
      syncControls();
      setView(state.view);
      renderAll();
    });
  }

  async function fetchDashboardPayload() {
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);

    try {
      const requestUrl = new URL("./data/contracts-dashboard.json", window.location.href);
      requestUrl.searchParams.set("ts", String(Date.now()));

      const response = await fetch(requestUrl.toString(), {
        cache: "no-store",
        credentials: "same-origin",
        mode: "same-origin",
        redirect: "error",
        signal: controller.signal,
        headers: {
          Accept: "application/json",
        },
      });

      if (!response.ok) {
        throw new Error("Informações indisponíveis.");
      }

      const payload = await response.json();
      if (!validatePayloadShape(payload)) {
        throw new Error("Dados inválidos.");
      }

      return payload;
    } finally {
      window.clearTimeout(timeoutId);
    }
  }

  async function bootstrap() {
    const initialState = readUrlState();
    state.view = initialState.view;
    state.filters = { ...DEFAULT_FILTERS, ...initialState.filters };

    const payload = await fetchDashboardPayload();
    state.payload = preprocessPayload(payload);
    state.filteredCache.clear();

    setLabeledOptions(elements.scopeSelect, ["atuais", "todos"], "scope");
    setPlainOptions(elements.organizationSelect, state.payload.filters.organizations || [], "Todos os órgãos");
    setPlainOptions(elements.administrationSelect, state.payload.filters.administrations || [], "Todas as gestões");
    setLabeledOptions(elements.vigencySelect, state.payload.filters.vigencyStates || Object.keys(LABELS.vigency), "vigency");
    setLabeledOptions(elements.managementSelect, state.payload.filters.managementStates || Object.keys(LABELS.management), "management");
    setLabeledOptions(elements.sourceSelect, state.payload.filters.sourceStates || Object.keys(LABELS.source), "source");
    setLabeledOptions(elements.criticalitySelect, Object.keys(LABELS.criticality), "criticality");
    setLabeledOptions(elements.sortSelect, Object.keys(LABELS.sort), "sort");

    ensureValidFilterValue("scope", elements.scopeSelect);
    ensureValidFilterValue("organization", elements.organizationSelect);
    ensureValidFilterValue("administration", elements.administrationSelect);
    ensureValidFilterValue("vigency", elements.vigencySelect);
    ensureValidFilterValue("management", elements.managementSelect);
    ensureValidFilterValue("source", elements.sourceSelect);
    ensureValidFilterValue("criticality", elements.criticalitySelect);
    ensureValidFilterValue("sort", elements.sortSelect);

    syncControls();
    setView(state.view);
    bindEvents();
    renderAll();
  }

  function renderError(message) {
    elements.heroSummary.textContent = message;
    elements.heroCallout.textContent = "";
    setHtml(elements.summaryCards, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.methodNotes, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.statusGrid, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.coverageGrid, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.sourceGrid, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.deadlineGrid, `<div class="empty-state">${escapeHtml(message)}</div>`);
    elements.pendingDeadlineSummary.textContent = message;
    setHtml(elements.pendingDeadlineRecords, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.organizationSummary, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.recentMovements, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.insightList, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.priorityGroups, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.alertSummaryGrid, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.alertRecords, `<div class="empty-state">${escapeHtml(message)}</div>`);
    elements.reviewQueueSummary.textContent = message;
    setHtml(elements.reviewSummaryGrid, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.reviewQueue, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.resultInsightGrid, `<div class="empty-state">${escapeHtml(message)}</div>`);
    setHtml(elements.recordList, `<div class="empty-state">${escapeHtml(message)}</div>`);
    elements.loadMore.classList.add("hidden");
  }

  bootstrap().catch(() => {
    renderError("Informações indisponíveis.");
  });
})();
