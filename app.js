const API = "https://script.google.com/macros/s/AKfycbzXPPImpVD97UnMutd33Hh3BSA1HWKLeHLgMvhPp_rLJtmdYZd3p-LwxA3gCTaFcK0G/exec";

const WEEK_DAYS = [
  { key: "Lunes", label: "Lunes", short: "Lun", offset: 0 },
  { key: "Martes", label: "Martes", short: "Mar", offset: 1 },
  { key: "Miercoles", label: "Miércoles", short: "Mie", offset: 2 },
  { key: "Jueves", label: "Jueves", short: "Jue", offset: 3 },
  { key: "Viernes", label: "Viernes", short: "Vie", offset: 4 },
  { key: "Sabado", label: "Sábado", short: "Sab", offset: 5 }
];
const DEFAULT_MEETING_DAY_KEYS = ["Martes", "Miercoles", "Jueves", "Sabado"];
const MEETING_DAYS_STORAGE_KEY = "discipulado_meeting_days_v1";
const WEEKDAY_NUMBER_BY_KEY = Object.freeze(
  WEEK_DAYS.reduce((acc, day) => {
    acc[day.key] = day.offset + 1;
    return acc;
  }, {})
);
const DAY_KEY_BY_NORMALIZED = Object.freeze(
  WEEK_DAYS.reduce((acc, day) => {
    acc[normalizeSearchText(day.key)] = day.key;
    return acc;
  }, {})
);
const SUBMISSION_CUTOFF_HOUR = 14;
const SUBMISSION_CUTOFF_LABEL = formatHourLabel(SUBMISSION_CUTOFF_HOUR);

const STATUS_META = {
  CONECTO: {
    label: "CLASE",
    chip: "CLASE",
    className: "status-connected",
    score: 3
  },
  ENVIADO: {
    label: "Enviado",
    chip: "Enviado",
    className: "status-sent",
    score: 1
  },
  NO: {
    label: "Fallido",
    chip: "Fallido",
    className: "status-missed",
    score: 0
  },
  PENDIENTE: {
    label: "Pendiente",
    chip: "Pendiente",
    className: "status-pending",
    score: 0
  }
};

const ATTENDANCE_META = {
  ASISTIO: {
    label: "Asistió",
    badge: "★",
    className: "attendance-corner attendance-corner-yes"
  },
  NO_ASISTIO: {
    label: "No asistió",
    badge: "✕",
    className: "attendance-corner attendance-corner-no"
  }
};

const state = {
  view: "discipulos",
  genero: "M",
  weekStart: getStartOfWeek(new Date()),
  rankingCollapsed: true,
  discipleStatsCollapsed: false,
  monthlyGuestsCollapsed: false,
  monthlyGuestsMonthKey: "",
  monthlyGuestsSaturdayISO: "",
  monthlyGuestsMemberId: "",
  memberSearchQuery: "",
  meetingDayKeys: [...DEFAULT_MEETING_DAY_KEYS],
  meetingDayNumbers: new Set(DEFAULT_MEETING_DAY_KEYS.map((key) => WEEKDAY_NUMBER_BY_KEY[key]).filter(Number.isFinite)),
  cache: {
    miembros: [],
    devos: []
  },
  cacheReady: false,
  attendanceOverrides: {},
  modal: {
    miembroId: "",
    fechaISO: "",
    estadoActual: "PENDIENTE",
    estadoSeleccionado: "",
    asistenciaActual: "",
    asistenciaSeleccionada: "",
    registroExistente: false,
    esReunion: false,
    esSabado: false,
    sabadoInvitados: 0
  },
  deleteModal: {
    miembroId: "",
    nombre: ""
  },
  loadingCount: 0,
  cutoffTickerId: null
};

const el = {};

document.addEventListener("DOMContentLoaded", init);

function init() {
  cacheElements();
  setupMeetingDaysControls();
  bindEvents();
  startCutoffTicker();
  cargarVista("discipulos");
}

function cacheElements() {
  el.contenido = document.getElementById("contenido");
  el.navDiscipulos = document.getElementById("navDiscipulos");
  el.navDiscipulas = document.getElementById("navDiscipulas");
  el.navRegistro = document.getElementById("navRegistro");
  el.globalLoader = document.getElementById("globalLoader");
  el.loaderText = document.getElementById("loaderText");
  el.markModal = document.getElementById("markModal");
  el.markModalTitle = document.getElementById("markModalTitle");
  el.markModalContext = document.getElementById("markModalContext");
  el.deleteMemberModal = document.getElementById("deleteMemberModal");
  el.deleteMemberModalTitle = document.getElementById("deleteMemberModalTitle");
  el.deleteMemberModalContext = document.getElementById("deleteMemberModalContext");
  el.deleteMemberConfirmBtn = document.getElementById("deleteMemberConfirmBtn");
  el.modalAttendanceWrap = document.getElementById("modalAttendanceWrap");
  el.modalGuestWrap = document.getElementById("modalGuestWrap");
  el.modalGuestInput = document.getElementById("modalGuestInput");
  el.modalSubmitBtn = document.getElementById("modalSubmitBtn");
  el.meetingDaysOptions = document.getElementById("meetingDaysOptions");
  el.meetingDaysSummary = document.getElementById("meetingDaysSummary");
  el.toastHost = document.getElementById("toastHost");
}

function bindEvents() {
  el.contenido.addEventListener("click", onContentClick);
  el.contenido.addEventListener("input", onContentInput);
  el.contenido.addEventListener("change", onContentInput);
  el.meetingDaysOptions?.addEventListener("change", onMeetingDaysInput);

  el.markModal.addEventListener("click", (event) => {
    if (event.target === el.markModal) {
      closeMarkModal();
      return;
    }

    const trigger = event.target.closest("[data-action]");
    if (!trigger) {
      return;
    }

    const action = trigger.dataset.action;
    if (action === "close-modal") {
      closeMarkModal();
      return;
    }

    if (action === "select-mark-state") {
      seleccionarEstadoModal(trigger.dataset.state);
      return;
    }

    if (action === "select-mark-attendance") {
      seleccionarAsistenciaModal(trigger.dataset.attendance);
      return;
    }

    if (action === "submit-mark") {
      guardarMarcaDesdeModal();
    }
  });

  el.deleteMemberModal.addEventListener("click", (event) => {
    if (event.target === el.deleteMemberModal) {
      closeDeleteMemberModal();
      return;
    }

    const trigger = event.target.closest("[data-action]");
    if (!trigger) {
      return;
    }

    const action = trigger.dataset.action;
    if (action === "close-delete-member-modal") {
      closeDeleteMemberModal();
      return;
    }

    if (action === "confirm-delete-member") {
      eliminarMiembroFrontend(state.deleteModal.miembroId);
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeMarkModal();
      closeDeleteMemberModal();
    }
  });
}

function setupMeetingDaysControls() {
  const persisted = readMeetingDaysFromStorage();
  applyMeetingDayKeys(persisted, false);
  renderMeetingDaysOptions();
}

function onMeetingDaysInput(event) {
  const checkbox = event.target.closest("[data-action='toggle-meeting-day']");
  if (!checkbox) {
    return;
  }

  const selectedKeys = Array.from(
    el.meetingDaysOptions?.querySelectorAll("[data-action='toggle-meeting-day']:checked") || []
  ).map((input) => input.dataset.dayKey);

  applyMeetingDayKeys(selectedKeys, true);
  closeMarkModal();

  if (state.view !== "registro") {
    renderDashboardView(false);
  }

  showToast("Días de reunión actualizados.", "success");
}

function onContentInput(event) {
  const monthFilter = event.target.closest("[data-action='filter-monthly-guests-month']");
  if (monthFilter) {
    state.monthlyGuestsMonthKey = normalizeMonthKey(monthFilter.value);
    state.monthlyGuestsSaturdayISO = "";
    state.monthlyGuestsMemberId = "";
    renderDashboardView(false);
    return;
  }

  const saturdayFilter = event.target.closest("[data-action='filter-monthly-guests-saturday']");
  if (saturdayFilter) {
    state.monthlyGuestsSaturdayISO = String(saturdayFilter.value || "").trim();
    state.monthlyGuestsMemberId = "";
    renderDashboardView(false);
    return;
  }

  const memberFilter = event.target.closest("[data-action='filter-monthly-guests-member']");
  if (memberFilter) {
    state.monthlyGuestsMemberId = String(memberFilter.value || "").trim();
    renderDashboardView(false);
    return;
  }

  const target = event.target.closest("[data-action='filter-members']");
  if (!target) {
    return;
  }

  state.memberSearchQuery = String(target.value || "").trim();
  applyCalendarFilter();
}

window.cargarVista = cargarVista;

async function cargarVista(vista) {
  state.view = vista;
  setActiveNav(vista);
  closeMarkModal();
  closeDeleteMemberModal();

  if (vista === "registro") {
    renderRegistroView();
    return;
  }

  state.genero = vista === "discipulas" ? "F" : "M";
  await renderDashboardView(!state.cacheReady);
}

function setActiveNav(vista) {
  [el.navDiscipulos, el.navDiscipulas, el.navRegistro].forEach((node) => node?.classList.remove("is-active"));

  if (vista === "discipulos") {
    el.navDiscipulos?.classList.add("is-active");
  } else if (vista === "discipulas") {
    el.navDiscipulas?.classList.add("is-active");
  } else {
    el.navRegistro?.classList.add("is-active");
  }
}

function renderRegistroView() {
  el.contenido.innerHTML = `
    <section class="form-shell">
      <header class="section-head">
        <h3>Registrar miembro</h3>
        <p>Agrega nuevos integrantes al ministerio y déjalos listos para seguimiento.</p>
      </header>

      <div class="form-grid">
        <label class="field-wrap">
          <span>Nombre completo</span>
          <input id="nombre" type="text" maxlength="80" placeholder="Ej. Juan Pérez" autocomplete="off">
        </label>

        <label class="field-wrap">
          <span>Género</span>
          <select id="genero">
            <option value="M">Masculino</option>
            <option value="F">Femenino</option>
          </select>
        </label>
      </div>

      <button class="primary-btn" type="button" data-action="save-member">Guardar miembro</button>
    </section>
  `;
}

async function renderDashboardView(forceReload) {
  const shouldShowSkeleton = forceReload || !el.contenido?.querySelector(".dashboard-wrap");
  if (shouldShowSkeleton) {
    renderDashboardSkeleton();
  }

  try {
    if (forceReload || !state.cacheReady) {
      await cargarDatosRemotos();
    }

    const html = buildDashboardMarkup();
    el.contenido.innerHTML = html;
    applyCalendarFilter();
  } catch (error) {
    el.contenido.innerHTML = `
      <div class="error-panel">
        <h3>No pudimos cargar el panel</h3>
        <p>${escapeHtml(error.message)}</p>
        <button class="primary-btn" type="button" data-action="refresh-view">Reintentar</button>
      </div>
    `;
  }
}

function renderDashboardSkeleton() {
  el.contenido.innerHTML = `
    <div class="skeleton-wrap">
      <div class="skeleton skeleton-lg"></div>
      <div class="skeleton-row">
        <div class="skeleton"></div>
        <div class="skeleton"></div>
        <div class="skeleton"></div>
        <div class="skeleton"></div>
      </div>
      <div class="skeleton skeleton-xl"></div>
      <div class="skeleton skeleton-xl"></div>
    </div>
  `;
}

async function cargarDatosRemotos() {
  await runWithLoader(async () => {
    const [miembrosRaw, devosRaw] = await Promise.all([
      apiRequestJSON(`${API}?tipo=miembros`),
      apiRequestJSON(`${API}?tipo=devos`)
    ]);

    state.cache.miembros = normalizeMiembros(miembrosRaw);
    state.cache.devos = normalizeDevos(devosRaw);
    state.cacheReady = true;
  }, "Sincronizando información...");
}

function buildDashboardMarkup() {
  const miembros = state.cache.miembros.filter((m) => m.genero === state.genero);
  const semana = getWeekDays(state.weekStart);
  const semanaSet = new Set(semana.map((d) => d.iso));
  const miembroIds = new Set(miembros.map((m) => m.id));

  const weekStatusMapBase = buildLatestStatusMap(state.cache.devos, miembroIds, semanaSet);
  const weekStatusMap = buildWeekStatusMapWithCutoff(miembros, semana, weekStatusMapBase);
  const totalStatusMap = buildLatestStatusMap(state.cache.devos, miembroIds);
  const metrics = buildMemberMetrics(miembros, weekStatusMap, totalStatusMap);
  const metricsById = new Map(metrics.map((item) => [item.id, item]));

  const totalSlots = miembros.length * WEEK_DAYS.length;
  const markedSlots = metrics.reduce((sum, m) => sum + m.weekMarked, 0);
  const totalConecto = metrics.reduce((sum, m) => sum + m.weekConecto, 0);
  const totalEnviado = metrics.reduce((sum, m) => sum + m.weekEnviado, 0);
  const totalNo = metrics.reduce((sum, m) => sum + m.weekNo, 0);
  const totalInvitadosSabado = metrics.reduce((sum, m) => sum + m.weekInvitados, 0);
  const progressPct = totalSlots > 0 ? Math.round((markedSlots / totalSlots) * 100) : 0;

  const ranking = [...metrics].sort((a, b) => {
    if (b.weekEnviado !== a.weekEnviado) return b.weekEnviado - a.weekEnviado;
    if (b.weekInvitados !== a.weekInvitados) return b.weekInvitados - a.weekInvitados;
    if (b.weekConecto !== a.weekConecto) return b.weekConecto - a.weekConecto;
    return a.nombre.localeCompare(b.nombre, "es");
  });

  const podiumTop = ranking.slice(0, 3);
  const headerTitle = state.genero === "F" ? "Discípulas" : "Discípulos";

  const calendarRows = miembros.length
    ? miembros.map((miembro) => {
      const memberMetric = metricsById.get(miembro.id) || buildEmptyMetrics(miembro);
      const memberProgress = Math.round((memberMetric.weekMarked / WEEK_DAYS.length) * 100);

      const dayCells = semana.map((day) => {
        const key = `${miembro.id}|${day.iso}`;
        const devoWeek = weekStatusMap.get(key);
        const devoRaw = totalStatusMap.get(key);
        const devo = devoWeek || devoRaw;
        const status = devo?.estado || "PENDIENTE";
        const meta = STATUS_META[status] || STATUS_META.PENDIENTE;
        const invitados = clampNonNegativeInt(devo?.invitados);
        const asistencia = resolveAttendanceValue(devo?.asistencia, key);
        const attendanceMeta = ATTENDANCE_META[asistencia];
        const esReunion = isMeetingDayISO(day.iso);
        const lockedByCutoff = Boolean(devo?.lockedByCutoff);
        const lockClass = lockedByCutoff ? "day-cell-locked" : "";
        const lockAttrs = lockedByCutoff ? `disabled aria-disabled="true" data-locked="1"` : `data-locked="0"`;
        const title = lockedByCutoff ? `${meta.label} · Bloqueado por corte ${SUBMISSION_CUTOFF_LABEL}` : meta.label;
        const guestBadge = isSaturdayISO(day.iso) && invitados > 0
          ? `<span class="day-cell-guests">+${invitados} invitados</span>`
          : "";
        const attendanceBadge = esReunion && attendanceMeta
          ? `<span class="${attendanceMeta.className}" aria-label="${escapeHtml(attendanceMeta.label)}" title="${escapeHtml(attendanceMeta.label)}">${attendanceMeta.badge}</span>`
          : "";

        return `
          <button
            type="button"
            class="day-cell ${meta.className} ${lockClass}"
            data-action="open-mark-modal"
            data-miembro-id="${escapeHtml(miembro.id)}"
            data-fecha="${day.iso}"
            data-estado="${status}"
            data-invitados="${invitados}"
            data-asistencia="${asistencia}"
            data-day-label="${day.short}"
            ${lockAttrs}
            title="${escapeHtml(title)}"
          >
            ${attendanceBadge}
            <span class="day-cell-state">${meta.chip}</span>
            ${guestBadge}
            <span class="day-cell-date">${formatShortDate(day.date)}</span>
          </button>
        `;
      }).join("");

      return `
        <div class="calendar-row" data-member-name="${escapeHtml(normalizeSearchText(miembro.nombre))}">
          <div class="member-cell">
            <button
              type="button"
              class="member-delete-btn"
              data-action="delete-member"
              data-id="${escapeHtml(miembro.id)}"
              data-name="${escapeHtml(miembro.nombre)}"
              aria-label="Eliminar ${escapeHtml(miembro.nombre)}"
              title="Eliminar miembro"
            >
              🗑️
            </button>
            <strong>${escapeHtml(miembro.nombre)}</strong>
            <div class="member-micro">
              <span>${memberMetric.weekMarked}/${WEEK_DAYS.length} días</span>
              <span>${memberMetric.weekConecto} En clase</span>
            </div>
            <div class="mini-progress"><span style="width:${memberProgress}%"></span></div>
          </div>
          ${dayCells}
        </div>
      `;
    }).join("")
    : `<p class="empty-inline">No hay miembros registrados para este grupo.</p>`;

  const podiumLayout = [
    { position: 2, item: podiumTop[1] || null },
    { position: 1, item: podiumTop[0] || null },
    { position: 3, item: podiumTop[2] || null }
  ];

  const rankingRows = podiumTop.length
    ? `
      <div class="podium-grid">
        ${podiumLayout.map(({ position, item }) => {
          if (!item) {
            return `
              <article class="podium-slot place-${position} is-empty" aria-hidden="true">
                <div class="podium-name">-</div>
                <div class="podium-meta">Sin registro</div>
                <div class="podium-block"><span class="podium-num">${position}</span></div>
              </article>
            `;
          }

          return `
            <article class="podium-slot place-${position}">
              <div class="podium-name">${escapeHtml(item.nombre)}</div>
              <div class="podium-meta">${item.weekEnviado} enviados · ${item.weekInvitados} invitados</div>
              <div class="podium-block">
                <span class="podium-badge">
                  <span class="podium-num">${position}</span>
                  ${position === 1 ? `<span class="podium-trophy" aria-hidden="true">🏆</span>` : ""}
                </span>
              </div>
            </article>
          `;
        }).join("")}
      </div>
    `
    : `<p class="empty-inline">Sin datos para generar ranking.</p>`;

  const statsCards = `
    <article class="stat-card">
      <p>Miembros activos</p>
      <strong>${miembros.length}</strong>
    </article>
    <article class="stat-card">
      <p>En clase (semana)</p>
      <strong>${totalConecto}</strong>
    </article>
    <article class="stat-card">
      <p>Enviaron (semana)</p>
      <strong>${totalEnviado}</strong>
    </article>
    <article class="stat-card">
      <p>Progreso semanal</p>
      <strong>${progressPct}%</strong>
      <small>${markedSlots}/${totalSlots || 0} marcaciones</small>
    </article>
    <article class="stat-card">
      <p>Invitados sábado</p>
      <strong>${totalInvitadosSabado}</strong>
      <small>Total de la semana</small>
    </article>
  `;

  const discipleStats = metrics.length
    ? metrics.map((m) => {
      const weekPct = Math.round((m.weekMarked / WEEK_DAYS.length) * 100);
      const historicalPct = m.totalMarked > 0 ? Math.round((m.totalConecto / m.totalMarked) * 100) : 0;

      return `
        <article class="disciple-stat-card">
          <header>
            <strong>${escapeHtml(m.nombre)}</strong>
            <span>${m.scoreWeek} pts semana</span>
          </header>
          <p>En clase ${m.weekConecto} · Enviado ${m.weekEnviado} · Fallido ${m.weekNo}</p>
          <p>Invitados sábado: ${m.weekInvitados}</p>
          <div class="progress-line">
            <span style="width:${weekPct}%"></span>
          </div>
          <small>Semana: ${weekPct}% | Histórico (En clase): ${historicalPct}%</small>
        </article>
      `;
    }).join("")
    : `<p class="empty-inline">No hay estadísticas disponibles.</p>`;

  const monthlyGuestsReport = buildMonthlyGuestsReport(miembros, [...totalStatusMap.values()], state.weekStart);
  const monthlyGuestMonthOptions = monthlyGuestsReport.monthKeys
    .map((monthKey) => `<option value="${monthKey}" ${monthKey === monthlyGuestsReport.selectedMonthKey ? "selected" : ""}>${escapeHtml(formatMonthPretty(monthKey))}</option>`)
    .join("");
  const monthlyGuestSaturdayOptions = monthlyGuestsReport.saturdayOptions
    .map((option) => `<option value="${escapeHtml(option.value)}" ${option.value === monthlyGuestsReport.selectedSaturdayISO ? "selected" : ""}>${escapeHtml(option.label)}</option>`)
    .join("");
  const monthlyGuestMemberOptions = monthlyGuestsReport.memberOptions
    .map((member) => `<option value="${escapeHtml(member.id)}" ${member.id === monthlyGuestsReport.selectedMemberId ? "selected" : ""}>${escapeHtml(member.nombre)}</option>`)
    .join("");
  const monthlyGuestBars = monthlyGuestsReport.rows.length
    ? monthlyGuestsReport.rows.map((row) => {
      const pct = monthlyGuestsReport.maxInvitados > 0 ? Math.round((row.invitados / monthlyGuestsReport.maxInvitados) * 100) : 0;
      const barWidth = row.invitados > 0 ? Math.max(8, pct) : 0;
      const invitadoLabel = row.invitados === 1 ? "invitado" : "invitados";
      const pctLabel = monthlyGuestsReport.maxInvitados > 0 ? `${pct}% del mayor registro` : "Sin invitados en el período";

      return `
        <article class="monthly-guest-bar-card">
          <header>
            <strong>${escapeHtml(row.nombre)}</strong>
            <span>${row.invitados} ${invitadoLabel}</span>
          </header>
          <div class="monthly-guest-bar-track">
            <span style="width:${barWidth}%"></span>
          </div>
          <small>${pctLabel} · ${escapeHtml(monthlyGuestsReport.contextLabel)}</small>
        </article>
      `;
    }).join("")
    : `<p class="empty-inline">No hay invitados registrados para los sábados del período seleccionado.</p>`;

  return `
    <section class="dashboard-wrap">
      <header class="dashboard-head">
        <div>
          <h3>${headerTitle} · Calendario Lunes a Sábado</h3>
          <p>${formatWeekRangeLabel(state.weekStart, semana[semana.length - 1].date)}</p>
        </div>
        <div class="week-actions">
          <button type="button" class="ghost-btn" data-action="prev-week">Semana anterior</button>
          <button type="button" class="ghost-btn" data-action="go-current-week">Semana actual</button>
          <button type="button" class="ghost-btn" data-action="next-week">Semana siguiente</button>
        </div>
      </header>

      <section class="stats-grid">${statsCards}</section>

      <section class="ranking-panel ${state.rankingCollapsed ? "is-collapsed" : ""}">
        <button type="button" class="ranking-toggle" data-action="toggle-ranking" aria-expanded="${state.rankingCollapsed ? "false" : "true"}">
          <span class="ranking-toggle-title">Ranking espiritual · Top 3</span>
          <span class="ranking-toggle-arrow" aria-hidden="true">▾</span>
        </button>
        <div class="ranking-collapsible">
          <div class="ranking-list">${rankingRows}</div>
        </div>
      </section>

      <section class="disciple-stats-panel ${state.discipleStatsCollapsed ? "is-collapsed" : ""}">
        <button type="button" class="disciple-toggle" data-action="toggle-disciple-stats" aria-expanded="${state.discipleStatsCollapsed ? "false" : "true"}">
          <span class="disciple-toggle-title">Estadísticas por discípulo</span>
          <span class="disciple-toggle-arrow" aria-hidden="true">▾</span>
        </button>
        <div class="disciple-collapsible">
          <p class="disciple-subtitle">Dinámico por filtro y semana</p>
          <div class="disciple-stats-grid">${discipleStats}</div>
        </div>
      </section>

      <section class="monthly-guests-panel ${state.monthlyGuestsCollapsed ? "is-collapsed" : ""}">
        <button type="button" class="monthly-guests-toggle" data-action="toggle-monthly-guests" aria-expanded="${state.monthlyGuestsCollapsed ? "false" : "true"}">
          <span class="monthly-guests-toggle-title">Informe de invitados por mes</span>
          <span class="monthly-guests-toggle-arrow" aria-hidden="true">▾</span>
        </button>
        <div class="monthly-guests-collapsible">
          <div class="monthly-guests-tools">
            <label class="monthly-guests-filter">
              <span>Mes</span>
              <select data-action="filter-monthly-guests-month">${monthlyGuestMonthOptions}</select>
            </label>
            <label class="monthly-guests-filter">
              <span>Sábado</span>
              <select data-action="filter-monthly-guests-saturday">${monthlyGuestSaturdayOptions}</select>
            </label>
            <label class="monthly-guests-filter">
              <span>Discípulo</span>
              <select data-action="filter-monthly-guests-member">
                <option value="">Todos los que llevaron</option>
                ${monthlyGuestMemberOptions}
              </select>
            </label>
          </div>

          <div class="monthly-guests-kpi-row">
            <article class="monthly-guests-kpi">
              <p>${escapeHtml(monthlyGuestsReport.totalLabel)}</p>
              <strong>${monthlyGuestsReport.totalInvitados}</strong>
            </article>
            <article class="monthly-guests-kpi">
              <p>Discípulos con invitados</p>
              <strong>${monthlyGuestsReport.totalContribuyentes}</strong>
            </article>
            <article class="monthly-guests-kpi">
              <p>Sábados tomados</p>
              <strong>${monthlyGuestsReport.saturdayIsos.length}</strong>
            </article>
          </div>

          <p class="monthly-guests-subtitle">
            ${escapeHtml(monthlyGuestsReport.periodLabel)} · ${escapeHtml(monthlyGuestsReport.saturdayLabel)}
          </p>
          <div class="monthly-guests-bars">${monthlyGuestBars}</div>
        </div>
      </section>

      <section class="calendar-panel">
        <div class="calendar-tools">
          <label class="calendar-search-wrap" for="calendarMemberSearch">
            <span>Buscar miembro</span>
            <input
              id="calendarMemberSearch"
              type="search"
              data-action="filter-members"
              placeholder="Escribe un nombre..."
              autocomplete="off"
              value="${escapeHtml(state.memberSearchQuery)}"
            >
          </label>
        </div>
        <div class="calendar-header-row">
          <div class="calendar-head member">Discípulo</div>
          ${semana.map((day) => `<div class="calendar-head">${day.short}<small>${formatShortDate(day.date)}</small></div>`).join("")}
        </div>
        <div class="calendar-body">${calendarRows}</div>
        <p class="empty-inline calendar-filter-empty is-hidden" data-role="calendar-empty-filter">No hay miembros que coincidan con ese filtro.</p>
      </section>

      <p class="status-summary">Enviados: ${totalEnviado} · Fallidos: ${totalNo} · Invitados sábado: ${totalInvitadosSabado}</p>
    </section>
  `;
}

function onContentClick(event) {
  const trigger = event.target.closest("[data-action]");
  if (!trigger) {
    return;
  }

  const action = trigger.dataset.action;

  if (action === "save-member") {
    guardarMiembro();
    return;
  }

  if (action === "refresh-view") {
    if (state.view === "registro") {
      renderRegistroView();
    } else {
      renderDashboardView(true);
    }
    return;
  }

  if (action === "prev-week") {
    state.weekStart = addDays(state.weekStart, -7);
    renderDashboardView(false);
    return;
  }

  if (action === "next-week") {
    state.weekStart = addDays(state.weekStart, 7);
    renderDashboardView(false);
    return;
  }

  if (action === "go-current-week") {
    state.weekStart = getStartOfWeek(new Date());
    renderDashboardView(false);
    return;
  }

  if (action === "toggle-ranking") {
    state.rankingCollapsed = !state.rankingCollapsed;
    renderDashboardView(false);
    return;
  }

  if (action === "toggle-disciple-stats") {
    state.discipleStatsCollapsed = !state.discipleStatsCollapsed;
    renderDashboardView(false);
    return;
  }

  if (action === "toggle-monthly-guests") {
    state.monthlyGuestsCollapsed = !state.monthlyGuestsCollapsed;
    renderDashboardView(false);
    return;
  }

  if (action === "open-mark-modal") {
    if (trigger.dataset.locked === "1") {
      showToast(`Este horario quedó cerrado por corte (${SUBMISSION_CUTOFF_LABEL}).`, "warning");
      return;
    }

    openMarkModal(
      trigger.dataset.miembroId,
      trigger.dataset.fecha,
      trigger.dataset.estado || "PENDIENTE",
      trigger.dataset.invitados,
      trigger.dataset.asistencia,
      trigger.dataset.locked === "1"
    );
    return;
  }

  if (action === "delete-member") {
    const miembroId = String(trigger.dataset.id || "").trim();
    if (!miembroId) {
      return;
    }

    const member = state.cache.miembros.find((m) => m.id === miembroId);
    const nombre = String(trigger.dataset.name || member?.nombre || "este miembro");
    openDeleteMemberModal(miembroId, nombre);
  }
}

function applyCalendarFilter() {
  if (!el.contenido) {
    return;
  }

  const rows = el.contenido.querySelectorAll(".calendar-row[data-member-name]");
  if (!rows.length) {
    return;
  }

  const query = normalizeSearchText(state.memberSearchQuery);
  let visibleRows = 0;

  rows.forEach((row) => {
    const memberName = String(row.dataset.memberName || "");
    const isVisible = !query || memberName.includes(query);
    row.classList.toggle("is-hidden", !isVisible);
    if (isVisible) {
      visibleRows += 1;
    }
  });

  const empty = el.contenido.querySelector("[data-role='calendar-empty-filter']");
  if (empty) {
    empty.classList.toggle("is-hidden", visibleRows > 0);
  }
}

async function guardarMiembro() {
  const nombreInput = document.getElementById("nombre");
  const generoInput = document.getElementById("genero");

  if (!nombreInput || !generoInput) {
    showToast("No encontramos el formulario activo.", "error");
    return;
  }

  const nombre = String(nombreInput.value || "").trim();
  const genero = normalizeGenero(generoInput.value);

  if (!nombre) {
    showToast("Escribe el nombre antes de guardar.", "warning");
    return;
  }

  if (!genero) {
    showToast("Selecciona un género válido.", "warning");
    return;
  }

  const button = document.querySelector("[data-action='save-member']");
  if (button) {
    button.disabled = true;
  }

  try {
    const responseText = await runWithLoader(async () => {
      return apiRequest("POST", {
        tipo: "miembro",
        nombre,
        genero
      });
    }, "Guardando miembro...");

    if (!String(responseText || "").toUpperCase().includes("OK")) {
      throw new Error(`Respuesta inesperada: ${responseText || "(vacía)"}`);
    }

    nombreInput.value = "";
    generoInput.value = genero;
    showToast("Miembro guardado correctamente.", "success");
  } catch (error) {
    showToast(`No se pudo guardar: ${error.message}`, "error");
  } finally {
    if (button) {
      button.disabled = false;
    }
  }
}

function openDeleteMemberModal(miembroId, nombre) {
  if (!miembroId) {
    return;
  }

  state.deleteModal.miembroId = String(miembroId);
  state.deleteModal.nombre = String(nombre || "").trim() || "este miembro";

  el.deleteMemberModalTitle.textContent = "¿Eliminar este miembro y todos sus registros?";
  el.deleteMemberModalContext.textContent = `Se eliminará a ${state.deleteModal.nombre} junto con su historial de devocionales. Esta acción no se puede deshacer.`;
  el.deleteMemberModal.classList.remove("hidden");
}

function closeDeleteMemberModal() {
  state.deleteModal.miembroId = "";
  state.deleteModal.nombre = "";
  if (el.deleteMemberConfirmBtn) {
    el.deleteMemberConfirmBtn.disabled = false;
  }
  el.deleteMemberModal.classList.add("hidden");
}

async function eliminarMiembroFrontend(miembroId) {
  const id = String(miembroId || "").trim();
  if (!id) {
    return;
  }

  closeDeleteMemberModal();

  try {
    const responseText = await runWithLoader(async () => {
      return apiRequest("POST", {
        tipo: "eliminar_miembro",
        miembro_id: id
      });
    }, "Eliminando miembro...");

    if (!String(responseText || "").toUpperCase().includes("OK")) {
      throw new Error(`Respuesta inesperada: ${responseText || "(vacía)"}`);
    }

    await runWithLoader(async () => {
      await cargarDatosRemotos();
    }, "Actualizando datos...");

    showToast("Miembro eliminado correctamente.", "success");
    renderDashboardView(false);
  } catch (error) {
    if (el.deleteMemberConfirmBtn) {
      el.deleteMemberConfirmBtn.disabled = false;
    }
    showToast("Error al eliminar: " + error.message, "error");
  }
}

function openMarkModal(miembroId, fechaISO, estadoActual, invitadosRaw = 0, asistenciaActual = "", lockedByCutoff = false) {
  if (!miembroId || !fechaISO) {
    return;
  }

  const member = state.cache.miembros.find((m) => m.id === miembroId);
  if (!member) {
    showToast("No se encontró el miembro seleccionado.", "error");
    return;
  }

  state.modal.miembroId = miembroId;
  state.modal.fechaISO = fechaISO;
  state.modal.esReunion = isMeetingDayISO(fechaISO);
  state.modal.esSabado = isSaturdayISO(fechaISO);

  const latest = findLatestDevoForDay(miembroId, fechaISO);
  const estadoBase = latest?.estado || estadoActual;
  const invitadosBase = latest?.invitados ?? invitadosRaw;
  const asistenciaBase = latest?.asistencia ?? asistenciaActual;
  const cutoffLock = lockedByCutoff || (normalizeEstado(estadoBase) === "PENDIENTE" && hasSubmissionCutoffPassed(fechaISO));

  if (cutoffLock) {
    showToast(`Este horario quedó cerrado por corte (${SUBMISSION_CUTOFF_LABEL}).`, "warning");
    return;
  }

  state.modal.estadoActual = normalizeEstado(estadoBase);
  state.modal.estadoSeleccionado = state.modal.estadoActual === "PENDIENTE" ? "" : state.modal.estadoActual;
  state.modal.asistenciaActual = normalizeAsistencia(asistenciaBase);
  state.modal.asistenciaSeleccionada = state.modal.asistenciaActual;
  state.modal.registroExistente = Boolean(latest);
  state.modal.sabadoInvitados = clampNonNegativeInt(invitadosBase);

  el.markModalTitle.textContent = `Marcar a ${member.nombre}`;
  el.markModalContext.textContent = `${formatDatePretty(fechaISO)} · Estado actual: ${STATUS_META[state.modal.estadoActual]?.label || "Pendiente"}`;
  el.modalAttendanceWrap.classList.toggle("hidden", !state.modal.esReunion);
  el.modalGuestWrap.classList.toggle("hidden", !state.modal.esSabado);
  el.modalGuestInput.value = String(state.modal.sabadoInvitados);

  el.markModal.querySelectorAll("[data-action='select-mark-state']").forEach((btn) => {
    const isCurrent = normalizeEstado(btn.dataset.state) === state.modal.estadoSeleccionado;
    btn.classList.toggle("is-selected", isCurrent);
  });

  el.markModal.querySelectorAll("[data-action='select-mark-attendance']").forEach((btn) => {
    const isCurrent = normalizeAsistencia(btn.dataset.attendance) === state.modal.asistenciaSeleccionada;
    btn.classList.toggle("is-selected", isCurrent);
  });
  refreshModalSubmitButton();

  el.markModal.classList.remove("hidden");
}

function closeMarkModal() {
  state.modal.miembroId = "";
  state.modal.fechaISO = "";
  state.modal.estadoActual = "PENDIENTE";
  state.modal.estadoSeleccionado = "";
  state.modal.asistenciaActual = "";
  state.modal.asistenciaSeleccionada = "";
  state.modal.registroExistente = false;
  state.modal.esReunion = false;
  state.modal.esSabado = false;
  state.modal.sabadoInvitados = 0;
  el.modalGuestInput.value = "0";
  refreshModalSubmitButton();
  el.markModal.classList.add("hidden");
}

function seleccionarEstadoModal(estado) {
  state.modal.estadoSeleccionado = normalizeEstado(estado);

  el.markModal.querySelectorAll("[data-action='select-mark-state']").forEach((btn) => {
    const isSelected = normalizeEstado(btn.dataset.state) === state.modal.estadoSeleccionado;
    btn.classList.toggle("is-selected", isSelected);
  });

  refreshModalSubmitButton();
}

function seleccionarAsistenciaModal(asistencia) {
  state.modal.asistenciaSeleccionada = normalizeAsistencia(asistencia);

  el.markModal.querySelectorAll("[data-action='select-mark-attendance']").forEach((btn) => {
    const isSelected = normalizeAsistencia(btn.dataset.attendance) === state.modal.asistenciaSeleccionada;
    btn.classList.toggle("is-selected", isSelected);
  });

  refreshModalSubmitButton();
}

function refreshModalSubmitButton() {
  if (!el.modalSubmitBtn) {
    return;
  }

  const hasSelectedState = normalizeEstado(state.modal.estadoSeleccionado) !== "PENDIENTE";
  const requiresAttendance = Boolean(state.modal.esReunion);
  const hasSelectedAttendance = normalizeAsistencia(state.modal.asistenciaSeleccionada) !== "";
  el.modalSubmitBtn.disabled = !hasSelectedState || (requiresAttendance && !hasSelectedAttendance);
  el.modalSubmitBtn.textContent = state.modal.registroExistente ? "Actualizar registro" : "Guardar registro";
}

async function guardarMarcaDesdeModal() {
  const miembroId = String(state.modal.miembroId || "").trim();
  const fechaISO = String(state.modal.fechaISO || "").trim();
  const estadoNormalizado = normalizeEstado(state.modal.estadoSeleccionado);
  const asistenciaNormalizada = normalizeAsistencia(state.modal.asistenciaSeleccionada);
  const isUpdate = Boolean(state.modal.registroExistente);
  const esReunion = Boolean(state.modal.esReunion);
  const esSabado = Boolean(state.modal.esSabado);
  const invitadosSabado = esSabado
    ? clampNonNegativeInt(el.modalGuestInput.value)
    : 0;

  if (!miembroId || !fechaISO || !estadoNormalizado || estadoNormalizado === "PENDIENTE") {
    showToast("No se pudo registrar la marca, datos incompletos.", "error");
    return;
  }
  if (esReunion && !asistenciaNormalizada) {
    showToast("Selecciona asistencia para la reunión.", "warning");
    return;
  }

  try {
    const accionTexto = isUpdate ? "Actualizando registro..." : "Guardando registro...";
    closeMarkModal();

    const responseText = await runWithLoader(async () => {
      const payload = {
        tipo: "devo",
        miembro_id: miembroId,
        fecha: fechaISO,
        estado: estadoNormalizado,
        upsert: true,
        actualizarSiExiste: true
      };

      if (esReunion) {
        payload.asistencia = asistenciaNormalizada;
        payload.asistencia_reunion = asistenciaNormalizada;
        payload.reunion = asistenciaNormalizada;
        payload.asistio = asistenciaNormalizada === "ASISTIO" ? "SI" : "NO";
      }

      if (esSabado) {
        payload.sabadoInvitados = invitadosSabado;
        payload.invitados = invitadosSabado;
      }

      return apiRequest("POST", payload);
    }, accionTexto);

    if (!String(responseText || "").toUpperCase().includes("OK")) {
      throw new Error(`Respuesta inesperada: ${responseText || "(vacía)"}`);
    }

    const overrideKey = `${miembroId}|${fechaISO}`;
    if (esReunion) {
      state.attendanceOverrides[overrideKey] = asistenciaNormalizada;
    } else if (state.attendanceOverrides[overrideKey]) {
      delete state.attendanceOverrides[overrideKey];
    }

    await runWithLoader(async () => {
      await cargarDatosRemotos();
    }, "Actualizando datos...");

    showToast(isUpdate ? "Registro actualizado correctamente." : "Registro guardado correctamente.", "success");
    renderDashboardView(false);

  } catch (error) {
    showToast(`No se pudo marcar: ${error.message}`, "error");
  }
}

async function runWithLoader(task, message) {
  showLoader(message);
  try {
    return await task();
  } finally {
    hideLoader();
  }
}

function showLoader(message) {
  state.loadingCount += 1;
  el.loaderText.textContent = message || "Cargando...";
  el.globalLoader.classList.remove("hidden");
}

function hideLoader() {
  state.loadingCount = Math.max(0, state.loadingCount - 1);
  if (state.loadingCount === 0) {
    el.globalLoader.classList.add("hidden");
  }
}

function showToast(text, type = "info") {
  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  toast.textContent = text;

  el.toastHost.appendChild(toast);

  requestAnimationFrame(() => {
    toast.classList.add("show");
  });

  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => toast.remove(), 250);
  }, 3200);
}

function findLatestDevoForDay(miembroId, fechaISO) {
  for (let i = state.cache.devos.length - 1; i >= 0; i -= 1) {
    const devo = state.cache.devos[i];
    if (devo.miembroId === miembroId && devo.fechaISO === fechaISO) {
      return devo;
    }
  }

  return null;
}

function clampNonNegativeInt(value) {
  const num = Number(value);
  if (!Number.isFinite(num) || num < 0) {
    return 0;
  }

  return Math.floor(num);
}

function startCutoffTicker() {
  if (state.cutoffTickerId) {
    return;
  }

  state.cutoffTickerId = setInterval(() => {
    if (state.view === "registro") {
      return;
    }

    if (!el.contenido?.querySelector(".dashboard-wrap")) {
      return;
    }

    if (el.markModal && !el.markModal.classList.contains("hidden")) {
      return;
    }

    renderDashboardView(false);
  }, 60000);
}

function readMeetingDaysFromStorage() {
  try {
    const raw = window.localStorage.getItem(MEETING_DAYS_STORAGE_KEY);
    if (raw == null) {
      return [...DEFAULT_MEETING_DAY_KEYS];
    }

    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [...DEFAULT_MEETING_DAY_KEYS];
    }

    const sanitized = sanitizeMeetingDayKeys(parsed);
    if (parsed.length > 0 && !sanitized.length) {
      return [...DEFAULT_MEETING_DAY_KEYS];
    }

    return sanitized;
  } catch {
    return [...DEFAULT_MEETING_DAY_KEYS];
  }
}

function saveMeetingDaysToStorage(keys) {
  try {
    window.localStorage.setItem(MEETING_DAYS_STORAGE_KEY, JSON.stringify(keys));
  } catch {
    // Si el navegador bloquea almacenamiento, mantenemos funcionamiento en memoria.
  }
}

function sanitizeMeetingDayKeys(values) {
  if (!Array.isArray(values)) {
    return [];
  }

  const keySet = new Set(
    values
      .map((value) => DAY_KEY_BY_NORMALIZED[normalizeSearchText(value)])
      .filter((key) => Boolean(key))
  );

  return WEEK_DAYS
    .map((day) => day.key)
    .filter((key) => keySet.has(key));
}

function applyMeetingDayKeys(keys, persist) {
  const normalizedKeys = sanitizeMeetingDayKeys(keys);
  state.meetingDayKeys = normalizedKeys;
  state.meetingDayNumbers = new Set(
    normalizedKeys
      .map((key) => WEEKDAY_NUMBER_BY_KEY[key])
      .filter(Number.isFinite)
  );

  syncMeetingDaysOptions();
  if (persist) {
    saveMeetingDaysToStorage(normalizedKeys);
  }
}

function renderMeetingDaysOptions() {
  if (!el.meetingDaysOptions) {
    return;
  }

  el.meetingDaysOptions.innerHTML = WEEK_DAYS.map((day) => `
    <label class="meeting-day-option">
      <input
        type="checkbox"
        data-action="toggle-meeting-day"
        data-day-key="${day.key}"
        ${state.meetingDayKeys.includes(day.key) ? "checked" : ""}
      >
      <span>${escapeHtml(day.label || day.key)}</span>
    </label>
  `).join("");

  refreshMeetingDaysSummary();
}

function syncMeetingDaysOptions() {
  if (el.meetingDaysOptions) {
    el.meetingDaysOptions
      .querySelectorAll("[data-action='toggle-meeting-day']")
      .forEach((input) => {
        const key = String(input.dataset.dayKey || "");
        input.checked = state.meetingDayKeys.includes(key);
      });
  }

  refreshMeetingDaysSummary();
}

function refreshMeetingDaysSummary() {
  if (!el.meetingDaysSummary) {
    return;
  }

  if (!state.meetingDayKeys.length) {
    el.meetingDaysSummary.textContent = "Activos: ninguno";
    return;
  }

  const selectedLabels = state.meetingDayKeys
    .map((key) => getWeekDayLabel(key))
    .join(", ");

  el.meetingDaysSummary.textContent = `Activos: ${selectedLabels}`;
}

function normalizeMiembros(rows) {
  if (!Array.isArray(rows)) {
    return [];
  }

  return rows
    .map((row) => {
      const id = getAny(row, ["id", "ID", "Id", "miembro_id"]);
      const nombre = getAny(row, ["nombre", "Nombre"]);
      const genero = getAny(row, ["genero", "Genero", "Género"]);

      return {
        id: String(id ?? "").trim(),
        nombre: String(nombre ?? "").trim() || "Sin nombre",
        genero: normalizeGenero(genero)
      };
    })
    .filter((row) => row.id && row.genero);
}

function normalizeDevos(rows) {
  if (!Array.isArray(rows)) {
    return [];
  }

  return rows
    .map((row, index) => {
      const id = getAny(row, ["id", "ID", "Id"]);
      const miembroId = getAny(row, ["miembro_id", "miembroId", "miembro", "Miembro"]);
      const fecha = getAny(row, ["fecha", "Fecha"]);
      const estado = getAny(row, ["estado", "Estado"]);
      const invitadosBase = getAny(row, ["invitados", "Invitados"]);
      const invitadosSabado = getAny(row, [
        "sabadoInvitados",
        "SabadoInvitados",
        "sabado_invitados",
        "sabadoInvitado",
        "SabadoInvitado"
      ]);
      const asistencia = getAny(row, [
        "asistencia",
        "Asistencia",
        "asistio",
        "Asistio",
        "asistencia_reunion",
        "asistenciaReunion",
        "reunion",
        "Reunion",
        "attended",
        "attendance"
      ]);

      return {
        id: String(id ?? "").trim() || `row_${index}`,
        miembroId: String(miembroId ?? "").trim(),
        fechaISO: toISODate(fecha),
        estado: normalizeEstado(estado),
        invitados: Math.max(clampNonNegativeInt(invitadosBase), clampNonNegativeInt(invitadosSabado)),
        asistencia: normalizeAsistencia(asistencia),
        order: index
      };
    })
    .filter((row) => row.miembroId && row.fechaISO);
}

function buildLatestStatusMap(devos, memberIds, allowedDates = null) {
  const map = new Map();

  devos.forEach((row) => {
    if (!memberIds.has(row.miembroId)) {
      return;
    }

    if (allowedDates && !allowedDates.has(row.fechaISO)) {
      return;
    }

    const key = `${row.miembroId}|${row.fechaISO}`;
    map.set(key, row);
  });

  return map;
}

function buildWeekStatusMapWithCutoff(miembros, semana, explicitMap) {
  const map = new Map(explicitMap);
  const now = new Date();

  miembros.forEach((miembro) => {
    semana.forEach((day) => {
      const key = `${miembro.id}|${day.iso}`;
      const cutoffPassed = hasSubmissionCutoffPassed(day.iso, now);
      const existing = map.get(key);

      if (existing) {
        if (cutoffPassed && normalizeEstado(existing.estado) === "PENDIENTE") {
          map.set(key, {
            ...existing,
            estado: "NO",
            invitados: clampNonNegativeInt(existing.invitados),
            asistencia: normalizeAsistencia(existing.asistencia),
            autoGenerated: true,
            lockedByCutoff: true
          });
        }

        return;
      }

      if (!cutoffPassed) {
        return;
      }

      map.set(key, {
        id: `auto_${key}`,
        miembroId: miembro.id,
        fechaISO: day.iso,
        estado: "NO",
        invitados: 0,
        asistencia: "",
        order: Number.MAX_SAFE_INTEGER,
        autoGenerated: true,
        lockedByCutoff: true
      });
    });
  });

  return map;
}

function buildMemberMetrics(miembros, weekStatusMap, totalStatusMap) {
  const metrics = miembros.map((m) => buildEmptyMetrics(m));
  const byId = new Map(metrics.map((m) => [m.id, m]));

  totalStatusMap.forEach((devo) => {
    const metric = byId.get(devo.miembroId);
    if (!metric) {
      return;
    }

    const status = STATUS_META[devo.estado] ? devo.estado : "PENDIENTE";
    if (status === "PENDIENTE") {
      return;
    }

    metric.totalMarked += 1;
    metric.scoreTotal += STATUS_META[status].score;

    if (status === "CONECTO") metric.totalConecto += 1;
    if (status === "ENVIADO") metric.totalEnviado += 1;
    if (status === "NO") metric.totalNo += 1;
    if (isSaturdayISO(devo.fechaISO)) metric.totalInvitados += clampNonNegativeInt(devo.invitados);

  });

  weekStatusMap.forEach((devo) => {
    const metric = byId.get(devo.miembroId);
    if (!metric) {
      return;
    }

    const status = STATUS_META[devo.estado] ? devo.estado : "PENDIENTE";
    if (status === "PENDIENTE") {
      return;
    }

    metric.weekMarked += 1;
    metric.scoreWeek += STATUS_META[status].score;

    if (status === "CONECTO") metric.weekConecto += 1;
    if (status === "ENVIADO") metric.weekEnviado += 1;
    if (status === "NO") metric.weekNo += 1;
    if (isSaturdayISO(devo.fechaISO)) metric.weekInvitados += clampNonNegativeInt(devo.invitados);
  });

  return metrics;
}

function buildMonthlyGuestsReport(miembros, latestDevos, fallbackDate = new Date()) {
  const memberMap = new Map(miembros.map((m) => [m.id, m]));
  const now = new Date();
  const todayISO = formatISODate(now);
  const currentMonthKey = formatMonthKey(now);
  const monthKeys = buildMonthKeysFromCurrent(now);

  const saturdayRows = latestDevos.filter((row) => memberMap.has(row.miembroId) && isSaturdayISO(row.fechaISO));
  let selectedMonthKey = normalizeMonthKey(state.monthlyGuestsMonthKey);
  if (!selectedMonthKey || !monthKeys.includes(selectedMonthKey)) {
    selectedMonthKey = normalizeMonthKey(formatMonthKey(fallbackDate)) || monthKeys[0] || currentMonthKey;
  }
  if (!monthKeys.includes(selectedMonthKey)) {
    selectedMonthKey = monthKeys[0] || currentMonthKey;
  }
  state.monthlyGuestsMonthKey = selectedMonthKey;

  const monthDate = toMonthDate(selectedMonthKey) || new Date();
  const saturdaysInMonth = getSaturdaysOfMonthISO(monthDate.getFullYear(), monthDate.getMonth());
  let availableSaturdays = saturdaysInMonth;
  if (selectedMonthKey === currentMonthKey) {
    availableSaturdays = saturdaysInMonth.filter((iso) => iso <= todayISO);
  } else if (selectedMonthKey > currentMonthKey) {
    availableSaturdays = [];
  }

  let selectedSaturdayISO = String(state.monthlyGuestsSaturdayISO || "").trim();
  if (selectedSaturdayISO && !availableSaturdays.includes(selectedSaturdayISO)) {
    selectedSaturdayISO = "";
  }
  state.monthlyGuestsSaturdayISO = selectedSaturdayISO;

  const consideredSaturdays = selectedSaturdayISO ? [selectedSaturdayISO] : availableSaturdays;
  const saturdaySet = new Set(consideredSaturdays);
  const totalsById = new Map();

  saturdayRows.forEach((row) => {
    if (!saturdaySet.has(row.fechaISO)) {
      return;
    }

    const invitados = clampNonNegativeInt(row.invitados);
    if (invitados <= 0) {
      return;
    }

    totalsById.set(row.miembroId, (totalsById.get(row.miembroId) || 0) + invitados);
  });

  const memberOptions = [...totalsById.entries()]
    .map(([id]) => memberMap.get(id))
    .filter((member) => Boolean(member))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, "es"));

  let selectedMemberId = String(state.monthlyGuestsMemberId || "").trim();
  if (selectedMemberId && !memberOptions.some((member) => member.id === selectedMemberId)) {
    selectedMemberId = "";
  }
  state.monthlyGuestsMemberId = selectedMemberId;

  let rows = [];
  if (selectedMemberId) {
    const selectedMember = memberMap.get(selectedMemberId);
    rows = [{
      id: selectedMemberId,
      nombre: selectedMember?.nombre || "Miembro",
      invitados: totalsById.get(selectedMemberId) || 0
    }];
  } else {
    rows = [...totalsById.entries()]
      .map(([id, invitados]) => ({
        id,
        nombre: memberMap.get(id)?.nombre || "Miembro",
        invitados
      }))
      .sort((a, b) => {
        if (b.invitados !== a.invitados) return b.invitados - a.invitados;
        return a.nombre.localeCompare(b.nombre, "es");
      });
  }

  const maxInvitados = rows.reduce((max, item) => Math.max(max, item.invitados), 0);
  const totalInvitados = rows.reduce((sum, item) => sum + item.invitados, 0);
  const totalContribuyentes = selectedMemberId
    ? (rows[0]?.invitados > 0 ? 1 : 0)
    : rows.length;

  const monthPretty = formatMonthPretty(selectedMonthKey);
  const lastAvailableSaturdayISO = availableSaturdays[availableSaturdays.length - 1] || "";
  const totalLabel = selectedSaturdayISO
    ? `Invitados del ${formatDatePretty(selectedSaturdayISO)}`
    : `Total de invitados este mes (${monthPretty})`;
  const periodLabel = selectedSaturdayISO
    ? `Filtro activo: ${formatDatePretty(selectedSaturdayISO)}`
    : (lastAvailableSaturdayISO
      ? `General al corte del ${formatDatePretty(lastAvailableSaturdayISO)}`
      : `General sin sábados pasados en ${monthPretty}`);
  const saturdayLabel = selectedSaturdayISO
    ? `Mes: ${monthPretty}`
    : `Sábados detectados: ${formatSaturdayDayList(availableSaturdays)}`;
  const contextLabel = selectedSaturdayISO
    ? formatDatePretty(selectedSaturdayISO)
    : (lastAvailableSaturdayISO ? `Corte ${formatDatePretty(lastAvailableSaturdayISO)}` : "Sin sábados pasados");
  const saturdayOptions = [
    {
      value: "",
      label: lastAvailableSaturdayISO
        ? `General del mes (corte ${formatDatePretty(lastAvailableSaturdayISO)})`
        : "General del mes"
    },
    ...availableSaturdays
      .slice()
      .reverse()
      .map((iso) => ({
        value: iso,
        label: formatDatePretty(iso)
      }))
  ];

  return {
    monthKeys,
    memberOptions,
    selectedMonthKey,
    selectedSaturdayISO,
    selectedMemberId,
    rows,
    maxInvitados,
    totalInvitados,
    totalContribuyentes,
    saturdayIsos: consideredSaturdays,
    saturdayOptions,
    totalLabel,
    periodLabel,
    saturdayLabel,
    contextLabel
  };
}

function buildEmptyMetrics(miembro) {
  return {
    id: miembro.id,
    nombre: miembro.nombre,
    weekConecto: 0,
    weekEnviado: 0,
    weekNo: 0,
    weekInvitados: 0,
    weekMarked: 0,
    totalConecto: 0,
    totalEnviado: 0,
    totalNo: 0,
    totalInvitados: 0,
    totalMarked: 0,
    scoreWeek: 0,
    scoreTotal: 0
  };
}

function getWeekDays(startDate) {
  return WEEK_DAYS.map((day) => {
    const date = addDays(startDate, day.offset);
    return {
      key: day.key,
      short: day.short,
      date,
      iso: formatISODate(date)
    };
  });
}

function getStartOfWeek(date) {
  const tmp = new Date(date);
  const day = tmp.getDay() || 7;
  tmp.setHours(0, 0, 0, 0);
  tmp.setDate(tmp.getDate() - day + 1);
  return tmp;
}

function addDays(date, amount) {
  const tmp = new Date(date);
  tmp.setDate(tmp.getDate() + amount);
  return tmp;
}

function normalizeMonthKey(value) {
  const raw = String(value || "").trim();
  return /^\d{4}-\d{2}$/.test(raw) ? raw : "";
}

function getDynamicMonthEndYear(currentYear) {
  const baseEndYear = 2031;
  if (currentYear < baseEndYear) {
    return baseEndYear;
  }

  const extensionBlocks = Math.floor((currentYear - baseEndYear) / 5) + 1;
  return baseEndYear + extensionBlocks * 5;
}

function buildMonthKeysFromCurrent(date = new Date()) {
  const source = new Date(date);
  if (Number.isNaN(source.getTime())) {
    return [];
  }

  const currentYear = source.getFullYear();
  const currentMonth = source.getMonth() + 1;
  const endYear = getDynamicMonthEndYear(currentYear);
  const list = [];

  for (let year = currentYear; year <= endYear; year += 1) {
    const monthStart = year === currentYear ? currentMonth : 1;
    for (let month = monthStart; month <= 12; month += 1) {
      list.push(`${year}-${String(month).padStart(2, "0")}`);
    }
  }

  return list;
}

function formatMonthKey(date) {
  const source = new Date(date);
  if (Number.isNaN(source.getTime())) {
    return "";
  }

  const y = source.getFullYear();
  const m = String(source.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function toMonthDate(monthKey) {
  const normalized = normalizeMonthKey(monthKey);
  if (!normalized) {
    return null;
  }

  const [y, m] = normalized.split("-").map((v) => Number(v));
  const date = new Date(y, m - 1, 1);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatMonthPretty(monthKey) {
  const monthDate = toMonthDate(monthKey);
  if (!monthDate) {
    return "Mes no disponible";
  }

  const formatted = new Intl.DateTimeFormat("es-CO", {
    month: "long",
    year: "numeric"
  }).format(monthDate);

  return capitalize(formatted);
}

function getSaturdaysOfMonthISO(year, monthIndex) {
  const list = [];
  const date = new Date(year, monthIndex, 1);

  while (date.getMonth() === monthIndex) {
    if (date.getDay() === 6) {
      list.push(formatISODate(date));
    }
    date.setDate(date.getDate() + 1);
  }

  return list;
}

function formatSaturdayDayList(saturdayIsos) {
  const dayNumbers = saturdayIsos
    .map((iso) => toDate(iso))
    .filter((date) => Boolean(date))
    .map((date) => String(date.getDate()));

  if (!dayNumbers.length) {
    return "ninguno";
  }

  if (dayNumbers.length === 1) {
    return dayNumbers[0];
  }

  if (dayNumbers.length === 2) {
    return `${dayNumbers[0]} y ${dayNumbers[1]}`;
  }

  return `${dayNumbers.slice(0, -1).join(", ")} y ${dayNumbers[dayNumbers.length - 1]}`;
}

function getWeekDayLabel(key) {
  const match = WEEK_DAYS.find((day) => day.key === key);
  return match?.label || key;
}

function formatHourLabel(hour24) {
  const normalizedHour = Number.isFinite(hour24) ? Math.min(23, Math.max(0, Math.floor(hour24))) : 0;
  const suffix = normalizedHour >= 12 ? "PM" : "AM";
  const hour12 = normalizedHour % 12 === 0 ? 12 : normalizedHour % 12;
  return `${hour12}:00 ${suffix}`;
}

function formatISODate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatShortDate(date) {
  const formatted = new Intl.DateTimeFormat("es-CO", {
    day: "numeric",
    month: "short"
  }).format(date);

  return formatted.replace(".", "");
}

function formatWeekRangeLabel(start, end) {
  const startFmt = new Intl.DateTimeFormat("es-CO", {
    weekday: "long",
    day: "numeric",
    month: "long"
  }).format(start);

  const endFmt = new Intl.DateTimeFormat("es-CO", {
    weekday: "long",
    day: "numeric",
    month: "long"
  }).format(end);

  const weekNumber = getWeekNumber(start);
  return `Semana ${weekNumber} · ${capitalize(startFmt)} al ${capitalize(endFmt)}`;
}

function formatDatePretty(iso) {
  const date = toDate(iso);
  if (!date) {
    return "Fecha no disponible";
  }

  const formatted = new Intl.DateTimeFormat("es-CO", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric"
  }).format(date);

  return capitalize(formatted);
}

function getWeekNumber(date) {
  const tmp = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = tmp.getUTCDay() || 7;
  tmp.setUTCDate(tmp.getUTCDate() + 4 - day);

  const yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1));
  return Math.ceil(((tmp - yearStart) / 86400000 + 1) / 7);
}

function normalizeGenero(value) {
  const raw = String(value || "").trim().toUpperCase();
  if (raw.startsWith("M")) return "M";
  if (raw.startsWith("F")) return "F";
  return "";
}

function normalizeEstado(value) {
  const raw = String(value || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  if (!raw) return "PENDIENTE";
  if (raw.includes("CONECT")) return "CONECTO";
  if (raw.includes("CLASE")) return "CONECTO";
  if (raw.includes("ENVI")) return "ENVIADO";
  if (raw.includes("FALL")) return "NO";
  if (raw.includes("NO")) return "NO";
  return "PENDIENTE";
}

function normalizeAsistencia(value) {
  const raw = String(value || "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  if (!raw) return "";
  if (raw === "ASISTIO" || raw === "ASISTE" || raw === "SI" || raw === "YES" || raw === "TRUE" || raw === "1" || raw.includes("ESTRELLA") || raw.includes("⭐")) {
    return "ASISTIO";
  }
  if (raw === "NO_ASISTIO" || raw === "NOASISTIO" || raw === "NO ASISTIO" || raw === "NO" || raw === "FALSE" || raw === "X" || raw === "✖" || raw === "0") {
    return "NO_ASISTIO";
  }
  if (raw.includes("NO ASIST")) return "NO_ASISTIO";
  if (raw.includes("ASIST")) return "ASISTIO";
  return "";
}

function resolveAttendanceValue(value, key) {
  const normalized = normalizeAsistencia(value);
  if (normalized) {
    return normalized;
  }

  return normalizeAsistencia(state.attendanceOverrides[key]);
}

function toISODate(value) {
  if (value == null || value === "") {
    return "";
  }

  if (typeof value === "number") {
    const fromNumber = new Date(value);
    return Number.isNaN(fromNumber.getTime()) ? "" : formatISODate(fromNumber);
  }

  const raw = String(value).trim();
  if (!raw) {
    return "";
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }

  if (/^\d{4}-\d{2}-\d{2}T/.test(raw)) {
    const parsedISO = new Date(raw);
    return Number.isNaN(parsedISO.getTime()) ? "" : formatISODate(parsedISO);
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return formatISODate(parsed);
  }

  return "";
}

function toDate(iso) {
  const raw = String(iso || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return null;
  }

  const date = new Date(`${raw}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function isSaturdayISO(iso) {
  const date = toDate(iso);
  if (!date) {
    return false;
  }

  return date.getDay() === 6;
}

function isMeetingDayISO(iso) {
  const date = toDate(iso);
  if (!date) {
    return false;
  }

  return state.meetingDayNumbers.has(date.getDay());
}

function hasSubmissionCutoffPassed(iso, now = new Date()) {
  const date = toDate(iso);
  if (!date) {
    return false;
  }

  if (isMeetingDayISO(iso)) {
    date.setHours(23, 59, 59, 999);
  } else {
    date.setHours(SUBMISSION_CUTOFF_HOUR, 0, 0, 0);
  }
  return now.getTime() >= date.getTime();
}

async function apiRequest(method, body) {
  let response;

  try {
    response = await fetch(API, {
      method,
      mode: "cors",
      credentials: "omit",
      cache: "no-store",
      redirect: "follow",
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
        Accept: "text/plain, application/json"
      },
      body: JSON.stringify(body)
    });
  } catch {
    throw new Error("No se pudo conectar con la API. Verifica despliegue y permisos del Apps Script.");
  }

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${text || "Error desconocido"}`);
  }

  return text;
}

async function apiRequestJSON(url) {
  let response;

  try {
    response = await fetch(url, {
      method: "GET",
      mode: "cors",
      credentials: "omit",
      cache: "no-store",
      redirect: "follow",
      headers: {
        Accept: "application/json, text/plain"
      }
    });
  } catch {
    throw new Error("No se pudo conectar con la API. Verifica despliegue y permisos del Apps Script.");
  }

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${text || "Error desconocido"}`);
  }

  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`La API no devolvió JSON válido: ${text.slice(0, 200)}`);
  }
}

function getAny(obj, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      return obj[key];
    }
  }

  return undefined;
}

function capitalize(value) {
  const text = String(value || "");
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function normalizeSearchText(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
