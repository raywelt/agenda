/*
 * Agenda Semana de la Educación
 * Lee data/AgendaPlana.xlsx directamente en el navegador (SheetJS) y lo
 * muestra como calendario, un día a la vez. Si el archivo se reemplaza,
 * la página lo detecta y se actualiza sola (revisión cada minuto).
 */
(function () {
  "use strict";

  const DATA_URL = "data/AgendaPlana.xlsx";
  const POLL_MS = 60 * 1000;
  const TIME_ZONE = "America/Bogota";
  const FILTERS = ["formato", "modalidad", "sede", "lugar"];
  const BREAK_RE = /^(receso|almuerzo|coffee\s*break|refrigerio|descanso|break)\b/i;

  const DOW = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  const MONTHS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
    "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

  const $ = (id) => document.getElementById(id);
  const els = {
    days: $("days"), dayHead: $("dayHead"), events: $("events"), status: $("status"),
    loader: $("loader"), fileInput: $("fileInput"), updated: $("updated"), toast: $("toast"),
    q: $("q"), clear: $("clearFilters"), toggle: $("filtersToggle"), filters: $("filters"),
    filtersCount: $("filtersCount"),
    selects: Object.fromEntries(FILTERS.map((f) => [f, $("f-" + f)])),
  };

  const state = {
    events: [],
    days: [],          // ["2026-10-19", ...]
    options: {},       // { formato: [{key,label}], ... }
    day: null,
    filters: { formato: "", modalidad: "", sede: "", lugar: "" },
    q: "",
    showPast: false,   // en el día actual, los eventos terminados se pliegan
    hash: null,
    loadedAt: null,
  };

  // ---------- Utilidades ----------
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const fold = (s) => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();

  const clean = (v) => {
    if (v === null || v === undefined) return "";
    return String(v).replace(/ /g, " ").replace(/[ \t]+/g, " ").trim();
  };

  const linkify = (html) => html.replace(/https?:\/\/[^\s<]+[^\s<.,;:)]/g,
    (u) => `<a href="${u}" target="_blank" rel="noopener">${u.length > 60 ? u.slice(0, 57) + "…" : u}</a>`);

  const cap = (s) => (s ? s.charAt(0).toUpperCase() + s.slice(1) : s);

  const pad = (n) => String(n).padStart(2, "0");

  const fmtTime = (min) => (min == null ? "" : `${Math.floor(min / 60)}:${pad(min % 60)}`);

  function hashBytes(buf) {
    // FNV-1a: suficiente para detectar que el archivo cambió.
    const bytes = new Uint8Array(buf);
    let h = 0x811c9dc5;
    for (let i = 0; i < bytes.length; i++) {
      h ^= bytes[i];
      h = Math.imul(h, 0x01000193);
    }
    return (h >>> 0).toString(16) + ":" + bytes.length;
  }

  // Fecha y hora "ahora" en Colombia (o simulada con ?hoy=2026-10-20T10:30).
  function nowInColombia() {
    const sim = new URLSearchParams(location.search).get("hoy");
    if (sim) {
      const m = sim.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{1,2}):(\d{2}))?/);
      if (m) return { date: `${m[1]}-${m[2]}-${m[3]}`, min: m[4] ? (+m[4]) * 60 + (+m[5]) : 0 };
    }
    const parts = Object.fromEntries(new Intl.DateTimeFormat("en-CA", {
      timeZone: TIME_ZONE, year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", hourCycle: "h23",
    }).formatToParts(new Date()).map((p) => [p.type, p.value]));
    return { date: `${parts.year}-${parts.month}-${parts.day}`, min: (+parts.hour) * 60 + (+parts.minute) };
  }

  function dayLabel(key, withYear) {
    const [y, m, d] = key.split("-").map(Number);
    const dt = new Date(Date.UTC(y, m - 1, d));
    return `${DOW[dt.getUTCDay()]} ${d} de ${MONTHS[m - 1]}` + (withYear ? ` de ${y}` : "");
  }

  // ---------- Lectura del Excel ----------
  function parseDate(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") {
      const dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 864e5);
      return `${dt.getUTCFullYear()}-${pad(dt.getUTCMonth() + 1)}-${pad(dt.getUTCDate())}`;
    }
    const s = clean(v);
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
    m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/); // dd/mm/aaaa
    if (m) return `${m[3].length === 2 ? "20" + m[3] : m[3]}-${pad(m[2])}-${pad(m[1])}`;
    return null;
  }

  function parseTime(v) {
    if (v === null || v === undefined || v === "") return null;
    if (typeof v === "number") {
      const frac = v - Math.floor(v);
      return Math.round(frac * 1440) % 1440;
    }
    const s = fold(v).replace(/\./g, "");
    const m = s.match(/(\d{1,2})(?::(\d{2}))?\s*(am|pm|a m|p m)?/);
    if (!m) return null;
    let h = +m[1];
    const min = m[2] ? +m[2] : 0;
    const ap = m[3] ? m[3].replace(/\s/g, "") : "";
    if (ap === "pm" && h < 12) h += 12;
    if (ap === "am" && h === 12) h = 0;
    return h * 60 + min;
  }

  function columnFor(header) {
    const h = fold(header);
    if (!h) return null;
    if (h === "dia" || h === "fecha") return "dia";
    if (h.startsWith("hora")) return /(final|fin|termin)/.test(h) ? "fin" : "inicio";
    if (h.startsWith("evento") || h.startsWith("actividad")) return "evento";
    if (h.startsWith("lugar")) return "lugar";
    if (h.startsWith("sede")) return "sede";
    if (h.startsWith("formato")) return "formato";
    if (h.startsWith("modalidad")) return "modalidad";
    return null;
  }

  function parseWorkbook(buf) {
    const wb = XLSX.read(buf, { type: "array" });
    const rows = [];
    wb.SheetNames.forEach((name, sheetIdx) => {
      const data = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
      const headerIdx = data.findIndex((r) => r && r.some((c) => columnFor(c) === "evento"));
      if (headerIdx < 0) return;
      const cols = {};
      data[headerIdx].forEach((c, i) => {
        const key = columnFor(c);
        if (key && !(key in cols)) cols[key] = i;
      });
      let lastDate = null;
      data.slice(headerIdx + 1).forEach((r, i) => {
        const get = (k) => (k in cols ? r[cols[k]] : null);
        const text = String(get("evento") ?? "").replace(/\r/g, "").trim();
        if (!text) return;
        const date = parseDate(get("dia")) || lastDate;
        if (!date) return;
        lastDate = date;
        const lines = text.split("\n").map((l) => l.replace(/[ \t ]+/g, " ").trim()).filter(Boolean);
        const ev = {
          id: `${sheetIdx}-${i}`,
          date,
          start: parseTime(get("inicio")),
          end: parseTime(get("fin")),
          title: lines[0].replace(/^"+|"+$/g, ""),
          details: lines.slice(1).join("\n"),
          lugar: cap(clean(get("lugar")).replace(/\s*\n\s*/g, " ")),
          sede: cap(clean(get("sede")).replace(/\s*\n\s*/g, " ")),
          formato: cap(clean(get("formato"))),
          modalidad: cap(clean(get("modalidad"))),
        };
        ev.isBreak = BREAK_RE.test(ev.title) && !ev.formato;
        rows.push(ev);
      });
    });
    return rows;
  }

  // Agrupa valores que solo difieren en mayúsculas/tildes/espacios
  // ("Por confirmar" = "Por Confirmar").
  function buildOptions(events) {
    const options = {};
    FILTERS.forEach((f) => {
      const map = new Map();
      events.forEach((ev) => {
        if (!ev[f]) return;
        const key = fold(ev[f]);
        ev[f + "Key"] = key;
        if (!map.has(key)) map.set(key, ev[f]);
      });
      options[f] = [...map.entries()]
        .map(([key, label]) => ({ key, label }))
        .sort((a, b) => a.label.localeCompare(b.label, "es"));
    });
    return options;
  }

  function sortEvents(a, b) {
    return a.date.localeCompare(b.date)
      || (a.start ?? 9999) - (b.start ?? 9999)
      || (a.end ?? 9999) - (b.end ?? 9999)
      || (a.sede || "").localeCompare(b.sede || "", "es")
      || (a.lugar || "").localeCompare(b.lugar || "", "es");
  }

  // ---------- Filtros ----------
  function hasActiveFilters() {
    return FILTERS.some((f) => state.filters[f]) || !!state.q;
  }

  function matches(ev) {
    if (ev.isBreak) {
      // Los recesos solo se muestran sin filtros, o filtrando por la sede donde ocurren.
      if (state.q || state.filters.formato || state.filters.modalidad || state.filters.lugar) return false;
      return !state.filters.sede || ev.sedeKey === state.filters.sede;
    }
    for (const f of FILTERS) {
      if (state.filters[f] && ev[f + "Key"] !== state.filters[f]) return false;
    }
    if (state.q) {
      const hay = fold([ev.title, ev.details, ev.lugar, ev.sede, ev.formato, ev.modalidad].join(" "));
      if (!fold(state.q).split(/\s+/).every((w) => hay.includes(w))) return false;
    }
    return true;
  }

  function countFor(day) {
    return state.events.filter((ev) => ev.date === day && !ev.isBreak && matches(ev)).length;
  }

  // ---------- Estado en la URL (#dia=...&sede=...) ----------
  function readHash() {
    const p = new URLSearchParams(location.hash.slice(1));
    FILTERS.forEach((f) => { state.filters[f] = p.get(f) || ""; });
    state.q = p.get("q") || "";
    return p.get("dia");
  }

  function writeHash() {
    const p = new URLSearchParams();
    if (state.day) p.set("dia", state.day);
    FILTERS.forEach((f) => { if (state.filters[f]) p.set(f, state.filters[f]); });
    if (state.q) p.set("q", state.q);
    history.replaceState(null, "", "#" + p.toString());
  }

  // Día por defecto: hoy si hay programación; si no, el próximo día con eventos;
  // si el evento ya pasó, el último día. Respeta los filtros que vengan en el enlace.
  function defaultDay() {
    const today = nowInColombia().date;
    const withEvents = state.days.filter((d) => countFor(d) > 0);
    const days = withEvents.length ? withEvents : state.days;
    return days.find((d) => d >= today) || days[days.length - 1] || null;
  }

  // ---------- Render ----------
  function renderSelects() {
    FILTERS.forEach((f) => {
      const sel = els.selects[f];
      const current = state.filters[f];
      sel.innerHTML = `<option value="">Todos</option>` + state.options[f]
        .map((o) => `<option value="${esc(o.key)}">${esc(o.label)}</option>`).join("");
      if (current && !state.options[f].some((o) => o.key === current)) state.filters[f] = "";
      sel.value = state.filters[f];
      sel.classList.toggle("is-active", !!state.filters[f]);
    });
    els.q.value = state.q;
    const n = FILTERS.filter((f) => state.filters[f]).length + (state.q ? 1 : 0);
    els.filtersCount.hidden = !n;
    els.filtersCount.textContent = n;
  }

  function renderDays() {
    const today = nowInColombia().date;
    els.days.innerHTML = state.days.map((d) => {
      const [, m, dd] = d.split("-").map(Number);
      const dow = dayLabel(d).split(" ")[0];
      const n = countFor(d);
      const cls = ["day-btn", d === today && "is-today", d < today && "is-past", !n && "is-empty"].filter(Boolean).join(" ");
      return `<button type="button" class="${cls}" data-day="${d}" aria-pressed="${d === state.day}"
        aria-label="${esc(dayLabel(d))}, ${n} eventos">
        <span class="dow">${dow.slice(0, 3)}</span><span class="num">${dd}</span>
        <span class="cnt">${n} ${n === 1 ? "evento" : "eventos"}</span></button>`;
    }).join("");
  }

  function chipClass(key) {
    let h = 0;
    for (const ch of key) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
    return "c" + (h % 5);
  }

  function renderEvent(ev, now) {
    const isToday = ev.date === now.date;
    const past = ev.date < now.date || (isToday && ev.end != null && ev.end <= now.min);
    const live = isToday && ev.start != null && ev.start <= now.min && (ev.end == null ? false : now.min < ev.end);
    const time = `${fmtTime(ev.start)}${ev.end != null ? " – " + fmtTime(ev.end) : ""}`;

    if (ev.isBreak) {
      return `<li class="break${past ? " is-past" : ""}"><span class="time">${time}</span>
        <span class="label">${esc(ev.title)}${ev.sede ? ` <small>· ${esc(ev.sede)}</small>` : ""}</span></li>`;
    }

    const state_ = live ? "En curso" : past ? "Finalizado" : ev._next ? "Próximo" : "";
    const meta = [
      ev.lugar && `<span><b>Lugar</b> ${esc(ev.lugar)}</span>`,
      ev.sede && `<span><b>Sede</b> ${esc(ev.sede)}</span>`,
    ].filter(Boolean).join("");
    const chips = [
      ev.formato && `<span class="chip ${chipClass(ev.formatoKey)}">${esc(ev.formato)}</span>`,
      ev.modalidad && `<span class="chip mod mod-${esc(ev.modalidadKey.replace(/\s+/g, "-"))}">${esc(ev.modalidad)}</span>`,
    ].filter(Boolean).join("");

    return `<li class="event${live ? " is-now" : ""}${past ? " is-past" : ""}${ev._next ? " is-next" : ""}">
      <div class="time">
        <span class="start">${fmtTime(ev.start) || "—"}</span>
        ${ev.end != null ? `<span class="end">a ${fmtTime(ev.end)}</span>` : ""}
        ${state_ ? `<span class="state">${state_}</span>` : ""}
      </div>
      <div class="body">
        ${chips ? `<div class="chips">${chips}</div>` : ""}
        <h2>${esc(ev.title)}</h2>
        ${ev.details ? `<p class="details">${linkify(esc(ev.details))}</p>` : ""}
        ${meta ? `<div class="meta">${meta}</div>` : ""}
      </div>
    </li>`;
  }

  function render() {
    renderSelects();
    renderDays();
    writeHash();

    const now = nowInColombia();
    const list = state.events.filter((ev) => ev.date === state.day && matches(ev));
    const real = list.filter((ev) => !ev.isBreak);

    // Marca los siguientes eventos que aún no empiezan (si es hoy).
    list.forEach((ev) => { ev._next = false; });
    if (state.day === now.date) {
      const upcoming = real.filter((ev) => ev.start != null && ev.start > now.min);
      const nextStart = upcoming.length ? upcoming[0].start : null;
      upcoming.forEach((ev) => { ev._next = ev.start === nextStart; });
    }

    els.dayHead.innerHTML = state.day
      ? `<h1>${esc(dayLabel(state.day))}</h1>
         <span class="summary">${real.length} ${real.length === 1 ? "evento" : "eventos"}${hasActiveFilters() ? " con los filtros aplicados" : ""}</span>`
      : "";

    // Si es hoy, primero lo que está en curso y lo que viene; lo terminado se pliega.
    const isPast = (ev) => ev.date === now.date && ev.end != null && ev.end <= now.min;
    const done = state.day === now.date ? list.filter(isPast) : [];
    const hidePast = done.length > 0 && !state.showPast && done.length < list.length;
    const doneReal = done.filter((ev) => !ev.isBreak).length;
    const toggle = hidePast || (state.showPast && done.length)
      ? `<li class="past-toggle"><button type="button" data-past>${hidePast
        ? `Ver ${doneReal} ${doneReal === 1 ? "evento finalizado" : "eventos finalizados"} de hoy`
        : "Ocultar eventos finalizados"}</button></li>`
      : "";
    els.events.innerHTML = toggle + list
      .filter((ev) => !(hidePast && isPast(ev)))
      .map((ev) => renderEvent(ev, now)).join("");

    if (!state.events.length) {
      els.status.textContent = "La agenda no tiene eventos todavía.";
    } else if (!real.length) {
      const alt = state.days.filter((d) => d !== state.day && countFor(d) > 0);
      els.status.innerHTML = `<p>No hay eventos este día con los filtros seleccionados.</p>` +
        (alt.length
          ? `<p>Hay coincidencias en:</p><div class="suggest">${alt.map((d) =>
            `<button type="button" data-day="${d}">${esc(dayLabel(d))} (${countFor(d)})</button>`).join("")}</div>`
          : `<div class="suggest"><button type="button" data-clear>Limpiar filtros</button></div>`);
    } else {
      els.status.textContent = "";
    }
  }

  function selectDay(d) {
    state.day = d;
    state.showPast = false;
    render();
    const active = els.days.querySelector('[aria-pressed="true"]');
    if (active) els.days.scrollTo({ left: active.offsetLeft - els.days.clientWidth / 2 + active.offsetWidth / 2, behavior: "smooth" });
    const controls = $("controls");
    const top = controls.getBoundingClientRect().top + window.scrollY;
    if (window.scrollY > top) window.scrollTo({ top, behavior: "smooth" });
  }

  function showToast(msg) {
    els.toast.textContent = msg;
    els.toast.hidden = false;
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => { els.toast.hidden = true; }, 4000);
  }

  // ---------- Carga ----------
  function applyWorkbook(buf, sourceLabel) {
    const events = parseWorkbook(buf).sort(sortEvents);
    state.events = events;
    state.options = buildOptions(events);
    state.days = [...new Set(events.map((ev) => ev.date))].sort();
    if (!state.days.includes(state.day)) state.day = defaultDay();
    state.loadedAt = new Date();
    els.loader.hidden = true;
    els.updated.textContent = `Datos: ${sourceLabel} · leído ${state.loadedAt.toLocaleString("es-CO", { dateStyle: "medium", timeStyle: "short" })}`;
    render();
  }

  async function fetchData() {
    const res = await fetch(`${DATA_URL}?t=${Date.now()}`, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.arrayBuffer();
  }

  async function load(initial) {
    try {
      const buf = await fetchData();
      const h = hashBytes(buf);
      if (h === state.hash) return;
      const firstLoad = state.hash === null;
      state.hash = h;
      applyWorkbook(buf, "AgendaPlana.xlsx");
      if (!firstLoad) showToast("Agenda actualizada");
    } catch (err) {
      if (initial) {
        console.warn("No se pudo cargar la agenda:", err);
        els.status.textContent = "";
        els.loader.hidden = false;
      }
    }
  }

  // ---------- Eventos de la interfaz ----------
  els.days.addEventListener("click", (e) => {
    const b = e.target.closest("[data-day]");
    if (b) selectDay(b.dataset.day);
  });

  els.events.addEventListener("click", (e) => {
    if (!e.target.closest("[data-past]")) return;
    state.showPast = !state.showPast;
    render();
  });

  els.status.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    if (b.dataset.day) selectDay(b.dataset.day);
    if ("clear" in b.dataset) els.clear.click();
  });

  FILTERS.forEach((f) => {
    els.selects[f].addEventListener("change", (e) => {
      state.filters[f] = e.target.value;
      render();
    });
  });

  let qTimer;
  els.q.addEventListener("input", (e) => {
    clearTimeout(qTimer);
    qTimer = setTimeout(() => { state.q = e.target.value.trim(); render(); }, 200);
  });

  els.clear.addEventListener("click", () => {
    FILTERS.forEach((f) => { state.filters[f] = ""; });
    state.q = "";
    render();
  });

  els.toggle.addEventListener("click", () => {
    const open = !els.filters.classList.contains("is-open");
    els.filters.classList.toggle("is-open", open);
    els.toggle.setAttribute("aria-expanded", String(open));
  });

  els.fileInput.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    applyWorkbook(await file.arrayBuffer(), file.name);
  });

  // Teclado: flechas izquierda/derecha cambian de día.
  document.addEventListener("keydown", (e) => {
    if (e.target.closest("input, select, textarea")) return;
    const i = state.days.indexOf(state.day);
    if (e.key === "ArrowRight" && i < state.days.length - 1) selectDay(state.days[i + 1]);
    if (e.key === "ArrowLeft" && i > 0) selectDay(state.days[i - 1]);
  });

  // ---------- Inicio ----------
  state.day = readHash();
  load(true);
  // Revisa cada minuto si AgendaPlana.xlsx cambió y refresca los estados "en curso".
  setInterval(() => {
    if (document.hidden) return;
    if (location.protocol !== "file:") load(false);
    if (state.events.length) render();
  }, POLL_MS);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden && location.protocol !== "file:") load(false);
  });
})();
