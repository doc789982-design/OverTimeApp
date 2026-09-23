/* ═══════════════════════════════════════════════════════════════
   OVERTIMETAB · веб-срез — логика интерфейса.
   Данные приходят из настоящего движка программы (см. server.py).
   ═══════════════════════════════════════════════════════════════ */

const MONTHS  = ["Январь","Февраль","Март","Апрель","Май","Июнь",
                 "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
const MONTHS_S= ["Янв","Фев","Мар","Апр","Май","Июн","Июл","Авг","Сен","Окт","Ноя","Дек"];
const WEEKDAYS= ["Пн","Вт","Ср","Чт","Пт","Сб","Вс"];

const state = {
  year:  new Date().getFullYear(),
  month: new Date().getMonth() + 1,
  view:  "month",            // month | year
  empId: 0,
  theme: localStorage.getItem("ot-theme") || "light",
  variant: localStorage.getItem("ot-variant") || "squares",
  search: "",
  groupSel: 0,               // 0 = все
  today: null,
  employees: [],
  groups: [],
  miniMax: 1,
};

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g,
  (c) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

async function api(path) {
  const r = await fetch(path);
  const j = await r.json();
  if (j.error) throw new Error(j.error);
  return j;
}

/* ═══════════════════════════════════════════════════════════════
   ОДОМЕТР: старое значение уезжает вверх, новое въезжает снизу.
   ═══════════════════════════════════════════════════════════════ */
function odo(el, val, html) {
  const prev = el.dataset.val;
  const set = (node) => html ? (node.innerHTML = val) : (node.textContent = val);
  if (prev === undefined) {                 // первый показ — без анимации
    el.dataset.val = val;
    el.innerHTML = '<span class="old"></span><span class="cur"></span>';
    set(el.querySelector(".cur"));
    return;
  }
  if (prev === val) return;
  el.dataset.val = val;
  const cur = el.querySelector(".cur"), old = el.querySelector(".old");
  html ? (old.innerHTML = prev) : (old.textContent = prev);
  set(cur);
  old.style.transform = "none";   old.style.opacity = "1";
  cur.style.transform = "translateY(105%)";
  void el.offsetWidth;            // фиксируем стартовое положение
  old.style.transform = "translateY(-105%)"; old.style.opacity = "0";
  cur.style.transform = "translateY(0)";
  setTimeout(() => { old.textContent = ""; old.style.cssText = "";
                     cur.style.cssText = ""; }, 260);
}

/* ═══════════════════════════════════════════════════════════════
   РЕЙКА ГРУПП
   ═══════════════════════════════════════════════════════════════ */
function renderGroups() {
  const rail = $("#groupRail");
  const mk = (id, label, name) => `
    <button class="g-dot ${state.groupSel === id ? "on" : ""}" data-g="${id}">
      ${esc(label)}<span class="lbl">${esc(name)}</span>
    </button>`;
  let html = mk(0, "Все", "Все сотрудники");
  for (const g of state.groups) {
    const words = g.name.split(/\s+/);
    const label = (words[0][0] + (words[1]?.[0] || "")).toUpperCase();
    html += mk(g.id, label, g.name);
  }
  rail.innerHTML = html;
  rail.querySelectorAll(".g-dot").forEach(b =>
    b.onclick = () => { state.groupSel = +b.dataset.g; renderGroups(); renderList(); });
}

/* ═══════════════════════════════════════════════════════════════
   СПИСОК СОТРУДНИКОВ + МИНИ-ГОД
   ═══════════════════════════════════════════════════════════════ */
function isFutureMonth(m) {
  const t = state.today;
  return state.year > t.y || (state.year === t.y && m > t.m);
}

function renderList() {
  const box = $("#empList");
  const q = state.search.trim().toLowerCase();
  const list = state.employees.filter(e =>
    (!state.groupSel || e.group_id === state.groupSel) &&
    (!q || e.fio.toLowerCase().includes(q) || e.position.toLowerCase().includes(q)));

  if (!list.length) {
    box.innerHTML = `<div class="empty" style="padding:30px 10px">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6">
        <circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
      Никого не найдено</div>`;
    return;
  }

  box.innerHTML = list.map(e => {
    let mini = "";
    if (state.variant !== "off") {
      const cells = e.mini.map((v, m) => {
        const fut = isFutureMonth(m + 1);
        const cur = state.today.y === state.year && state.today.m === m + 1;
        const picked = e.id === state.empId &&
          ((state.view === "month" && state.month === m + 1) || state.view === "year");
        const tip = fut
          ? `${MONTHS[m]} ${state.year} — ещё не наступил`
          : `${MONTHS[m]} ${state.year} — всего дней: ${v}`;
        if (state.variant === "squares") {
          const bg = fut ? "" :
            v === 0 ? `background:var(--zero)` :
            `background:rgba(3,116,181,${(0.16 + 0.74 * Math.min(1, v / state.miniMax)).toFixed(2)})`;
          return `<i class="m ${fut ? "future" : ""} ${cur ? "cur" : ""} ${picked ? "picked" : ""}"
                    style="${bg}" title="${esc(tip)}" data-m="${m + 1}"></i>`;
        }
        const h = fut || v === 0 ? "" :
          `height:${Math.max(4, Math.round(v / state.miniMax * 24))}px;opacity:${(0.5 + 0.5 * Math.min(1, v / state.miniMax)).toFixed(2)}`;
        return `<i class="m ${fut ? "future" : v === 0 ? "zero" : ""} ${cur ? "cur" : ""} ${picked ? "picked" : ""}"
                  title="${esc(tip)}" data-m="${m + 1}">${h ? `<i style="${h}"></i>` : "<i></i>"}</i>`;
      }).join("");
      mini = `<div class="mini ${state.variant === "squares" ? "sq" : "bars"}">${cells}</div>`;
    }
    return `
      <div class="emp ${e.id === state.empId ? "sel" : ""}" data-id="${e.id}">
        <div class="info">
          <div class="fio">${esc(e.fio)}</div>
          <div class="pos">${esc(e.position)}</div>
          <div class="normbar"><i style="width:${Math.min(100, Math.round(e.ratio * 88))}%"></i></div>
        </div>
        ${mini}
      </div>`;
  }).join("");

  box.querySelectorAll(".emp").forEach(el =>
    el.onclick = (ev) => {
      selectEmp(+el.dataset.id);
      const m = ev.target.closest(".m:not(.future)")?.dataset.m;
      if (m) gotoMonth(+m);
    });
}

function recalcMiniMax() {
  const t = state.today;
  let mx = 1;
  for (const e of state.employees)
    for (let m = 0; m < (t.y === state.year ? t.m : 12); m++)
      mx = Math.max(mx, e.mini[m]);
  state.miniMax = mx;
}

/* ═══════════════════════════════════════════════════════════════
   ВКЛАДКИ ПЕРИОДА
   ═══════════════════════════════════════════════════════════════ */
function renderTabs() {
  const t = state.today;
  $("#yearBtn").innerHTML = `${state.year}
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"
         stroke-linecap="round"><path d="m6 9 6 6 6-6"/></svg>`;
  let html = "";
  for (let m = 1; m <= 12; m++) {
    const on = state.view === "month" && state.month === m;
    const dot = t.y === state.year && t.m === m ? '<span class="dot"></span>' : "";
    html += `<button class="tab ${on ? "on" : ""}" data-m="${m}">${MONTHS_S[m - 1]}${dot}</button>`;
  }
  html += `<button class="tab year-tab ${state.view === "year" ? "on" : ""}" data-m="year">Год</button>`;
  $("#monthTabs").innerHTML = html;
  $("#monthTabs").querySelectorAll(".tab").forEach(b =>
    b.onclick = () => b.dataset.m === "year" ? showYear() : gotoMonth(+b.dataset.m));
}

/* ═══════════════════════════════════════════════════════════════
   КАЛЕНДАРЬ МЕСЯЦА
   ═══════════════════════════════════════════════════════════════ */
function renderCalendar(data) {
  const t = state.today;
  const todayIso = `${t.y}-${String(t.m).padStart(2, "0")}-${String(t.d).padStart(2, "0")}`;
  const wd = WEEKDAYS.map(w => `<div>${w}</div>`).join("");
  const cells = data.days.map(d => {
    const cls = ["day"];
    if (!d.in_month) cls.push("out");
    if (d.is_holiday) cls.push("hol");
    if (d.date === todayIso) cls.push("today");
    const nCls = ["n"];
    if (d.is_weekend || d.is_holiday) nCls.push("red");
    let flags = "";
    if (d.has_comp) flags += `<span class="badge comp" title="Компенсация (приказ)">В</span>`;
    if (d.status) flags += `<span class="badge st-${d.status}" title="${d.status === "О" ? "Отпуск" : "Больничный"}">${d.status}</span>`;
    const duties = d.duties.map(x =>
      `<span class="duty ${x.is_shift ? "shift" : ""}">${esc(x.text)}</span>`).join("");
    const pre = d.is_pre_holiday ? '<span class="pre" title="Предпраздничный день"></span>' : "";
    const [y, m, dd] = d.date.split("-");
    const tip = `${dd}.${m}.${y}` +
      (d.is_holiday ? " · праздник" : d.is_weekend ? " · выходной" : "") +
      (d.is_pre_holiday ? " · предпраздничный" : "") +
      (d.duties.length ? "\nДежурства: " + d.duties.map(x => x.text).join(", ") : "") +
      (d.has_comp ? "\nКомпенсация" : "") + (d.status ? "\nСтатус: " + d.status : "");
    return `<div class="${cls.join(" ")}" title="${esc(tip)}">
      <div class="top"><span class="${nCls.join(" ")}">${d.n}</span>
        <span class="flags">${flags}</span></div>
      <div class="duties">${duties}</div>${pre}
    </div>`;
  }).join("");
  $("#calendarWrap").className = "calendar-wrap";
  $("#calendarWrap").innerHTML = `
    <div class="cal-head">${wd}</div>
    <div class="cal-grid">${cells}</div>`;
}

/* ═══════════════════════════════════════════════════════════════
   ВИД «ГОД» — панорама двенадцати месяцев
   ═══════════════════════════════════════════════════════════════ */
function renderYearView(data) {
  const t = state.today;
  const todayIso = `${t.y}-${String(t.m).padStart(2, "0")}-${String(t.d).padStart(2, "0")}`;
  const hol = new Set(data.panorama.holidays);
  const blocks = data.panorama.months.map(mo => {
    const fut = state.year === t.y && mo.m > t.m;
    const cells = mo.weeks.flat().map(iso => {
      const dt = new Date(iso + "T00:00:00");
      const mins = data.panorama.per_day[iso] || 0;
      const inM = dt.getMonth() + 1 === mo.m;
      let style = "", cls = "dm-cell";
      if (hol.has(iso)) { cls += " hol"; style = ""; }
      else if (dt.getDay() === 0 || dt.getDay() === 6) cls += " wknd";
      if (mins > 0 && inM) {
        const a = 0.2 + 0.7 * Math.min(1, mins / 1440);
        style += `background:rgba(3,116,181,${a.toFixed(2)})`;
      }
      if (iso === todayIso) cls += " today";
      return `<i class="${cls}" style="${style}" title="${iso}${mins ? " · " + Math.round(mins / 60) + " ч." : ""}"></i>`;
    }).join("");
    return `<div class="ym ${fut ? "fut" : ""}" data-m="${mo.m}">
      <div class="ym-head"><span class="ym-name">${MONTHS[mo.m - 1]}</span>
        <span class="ym-total ${mo.total < 0 ? "neg" : ""}"
          title="${fut ? "Будущий месяц — дежурств ещё нет, прогноз без них" : ""}">${mo.total} дн.</span></div>
      <div class="dm">${cells}</div>
    </div>`;
  }).join("");
  $("#calendarWrap").className = "calendar-wrap";
  $("#calendarWrap").innerHTML = `
    <div class="year-view">
      <div class="year-title">
        <h2>${state.year} год</h2>
        <span class="res">остаток на конец года: <b>${data.mini[11]} дн.</b></span>
      </div>
      <div class="ym-grid">${blocks}</div>
    </div>`;
  $("#calendarWrap").querySelectorAll(".ym").forEach(el =>
    el.onclick = () => gotoMonth(+el.dataset.m));
}

/* ═══════════════════════════════════════════════════════════════
   ПАНЕЛЬ БАЛАНСОВ
   ═══════════════════════════════════════════════════════════════ */
function renderSummary(data) {
  const p = $("#summaryPanel");
  if (!data) {
    p.innerHTML = `<div class="empty">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6">
        <rect x="3" y="4" width="18" height="17" rx="3"/><path d="M8 2v4M16 2v4M3 9h18"/></svg>
      Выберите сотрудника</div>`;
    return;
  }
  const s = data.summary;
  const ratio = Math.min(1.4, Math.max(0, data.ratio ?? 0));
  const cards = s.cards.map(c => `
    <div class="bcard">
      <div class="bt">${esc(c.title)}</div>
      <div class="rows">
        <div class="r"><span>на начало</span><span class="v">${c.start}</span></div>
        <div class="r"><span>начислено</span><span class="v">${c.acc}</span></div>
        <div class="r"><span>компенсировано</span><span class="v">${c.comp}</span></div>
      </div>
      <div class="end">
        <span class="cap">остаток</span>
        <span class="v ${c.end_neg ? "neg" : ""}"><span class="odo" data-k="${esc(c.title)}"></span></span>
      </div>
    </div>`).join("");

  p.innerHTML = `
    <div class="sum-head">
      <div class="period">${esc(data.period)}</div>
      <div class="who">${esc(data.emp.fio)} · ${esc(data.emp.position)}</div>
      <div class="tags">
        ${s.is_shift_month ? '<span class="tag shift">сменный месяц</span>' : ""}
        <span class="tag">ночные: ${esc(s.night)}</span>
      </div>
    </div>
    <div class="total-card">
      <div class="cap">ВСЕГО ДНЕЙ</div>
      <div class="val"><span class="odo" id="odoTotal"></span> <span
        style="font-size:15px;font-weight:700">дн.</span></div>
      <div class="sub">остаток на конец месяца · ночные + сверх нормы + ДДО</div>
    </div>
    <div class="norm-row">
      <div class="line"><span>Норма <b>${esc(s.norm)}</b></span><span>Отработано <b>${esc(s.shift)}</b></span></div>
      <div class="normbar big"><i style="width:${Math.min(100, Math.round(ratio * 86))}%"></i></div>
    </div>
    ${cards}`;

  odo($("#odoTotal"), String(s.total_days));
  p.querySelectorAll(".bcard").forEach(el => {
    const title = el.querySelector(".bt").textContent;
    const c = s.cards.find(x => x.title === title);
    odo(el.querySelector(".odo"), c.end_plain);
  });
}

function renderYearSummary(data) {
  const p = $("#summaryPanel");
  const t = state.today;
  const rows = data.mini.map((v, i) => {
    const fut = state.year === t.y && i + 1 > t.m;
    return `<div class="r"><span>${MONTHS_S[i]}</span>
      <span class="v ${fut ? "fut" : ""}" ${fut ? 'title="Будущий месяц — прогноз"' : ""}>${v} дн.</span></div>`;
  }).join("");
  p.innerHTML = `
    <div class="sum-head">
      <div class="period">${state.year} год</div>
      <div class="who">${esc(data.emp.fio)}</div>
    </div>
    <div class="total-card">
      <div class="cap">ОСТАТОК НА КОНЕЦ ГОДА</div>
      <div class="val"><span class="odo" id="odoTotal"></span> <span
        style="font-size:15px;font-weight:700">дн.</span></div>
      <div class="sub">прогноз: будущие месяцы без дежурств</div>
    </div>
    <div class="bcard">
      <div class="bt">Всего дней по месяцам</div>
      <div class="rows">${rows}</div>
    </div>`;
  odo($("#odoTotal"), String(data.mini[11]));
}

/* ═══════════════════════════════════════════════════════════════
   ЗАГРУЗКА ДАННЫХ
   ═══════════════════════════════════════════════════════════════ */
let monthCache = null;

async function loadBootstrap() {
  const b = await api(`/api/bootstrap?year=${state.year}`);
  state.today = b.today;
  state.groups = b.groups;
  state.employees = b.employees;
  recalcMiniMax();
  renderGroups();
  renderList();
  renderTabs();
}

async function loadMonth() {
  const wrap = $("#calendarWrap");
  wrap.className = "calendar-wrap loading";
  wrap.innerHTML = `<div class="cal-head">${WEEKDAYS.map(w => `<div>${w}</div>`).join("")}</div>
    <div class="cal-grid">${"<div class=\"skeleton\"></div>".repeat(42)}</div>`;
  try {
    const d = await api(`/api/month?emp=${state.empId}&year=${state.year}&month=${state.month}`);
    monthCache = d;
    renderCalendar(d);
    renderSummary(d);
    renderTabs();
    updateTodayBtn();
  } catch (e) {
    wrap.className = "calendar-wrap";
    wrap.innerHTML = `<div class="empty">Ошибка: ${esc(e.message)}</div>`;
  }
}

async function loadYear() {
  const wrap = $("#calendarWrap");
  wrap.className = "calendar-wrap loading";
  wrap.innerHTML = `<div class="year-view"><div class="ym-grid">
    ${"<div class=\"skeleton\" style=\"height:150px\"></div>".repeat(12)}</div></div>`;
  try {
    const d = await api(`/api/year?emp=${state.empId}&year=${state.year}`);
    renderYearView(d);
    renderYearSummary(d);
    renderTabs();
    updateTodayBtn();
  } catch (e) {
    wrap.className = "calendar-wrap";
    wrap.innerHTML = `<div class="empty">Ошибка: ${esc(e.message)}</div>`;
  }
}

/* ═══════════════════════════════════════════════════════════════
   ДЕЙСТВИЯ
   ═══════════════════════════════════════════════════════════════ */
function selectEmp(id) {
  if (state.empId === id) return;
  state.empId = id;
  renderList();
  state.view === "year" ? loadYear() : loadMonth();
}

function gotoMonth(m) {
  state.month = m;
  state.view = "month";
  loadMonth();
}

function showYear() {
  state.view = "year";
  loadYear();
}

async function setYear(y) {
  if (state.year === y) return;
  state.year = y;
  $("#yearWrap")?.classList.remove("open");
  await loadBootstrap();          // мини-годы пересчитаются под новый год
  state.view === "year" ? loadYear() : loadMonth();
}

function updateTodayBtn() {
  const t = state.today;
  const far = state.year !== t.y ||
    (state.view === "year" ? false : state.month !== t.m);
  $("#todayBtn").classList.toggle("show", far);
}

/* ═══════════════════════════════════════════════════════════════
   ШАПКА: тема · вариант мини-года · меню года · поиск
   ═══════════════════════════════════════════════════════════════ */
function applyTheme() {
  document.documentElement.dataset.theme = state.theme;
  $("#iconMoon").style.display = state.theme === "light" ? "" : "none";
  $("#iconSun").style.display = state.theme === "dark" ? "" : "none";
}

function applyVariant() {
  localStorage.setItem("ot-variant", state.variant);
  document.querySelectorAll("#variantSeg button").forEach(b =>
    b.classList.toggle("on", b.dataset.v === state.variant));
  renderList();
}

/* ═══════════════════════════════════════════════════════════════
   СТАРТ
   ═══════════════════════════════════════════════════════════════ */
async function init() {
  applyTheme();
  applyVariant();
  try {
    await loadBootstrap();
  } catch (e) {
    document.body.innerHTML =
      `<div style="height:100vh;display:grid;place-items:center;text-align:center;
        font-family:Segoe UI,Arial,sans-serif;color:#6B7780;padding:20px">
        <div><b style="color:#2D3B45">Нет связи с сервером среза.</b><br><br>
        Откройте живое превью процесса «Веб-срез OVERTIMETAB» (порт 8081) —<br>
        интерфейсу нужен движок программы, который считает данные.</div></div>`;
    return;
  }
  if (state.employees.length) {
    state.empId = state.employees[0].id;
    renderList();
  }
  loadMonth();

  $("#themeBtn").onclick = () => {
    state.theme = state.theme === "light" ? "dark" : "light";
    localStorage.setItem("ot-theme", state.theme);
    applyTheme();
  };
  document.querySelectorAll("#variantSeg button").forEach(b =>
    b.onclick = () => { state.variant = b.dataset.v; applyVariant(); });
  $("#search").oninput = (e) => { state.search = e.target.value; renderList(); };
  $("#todayBtn").onclick = async () => {
    if (state.year !== state.today.y) await setYear(state.today.y);
    gotoMonth(state.today.m);
  };

  // меню года
  const yw = document.querySelector(".year-wrap");
  $("#yearBtn").onclick = (e) => {
    e.stopPropagation();
    const y0 = Math.min(...state.employees.map(x => +x.start_month.slice(0, 4)), state.today.y - 2);
    const menu = $("#yearMenu");
    menu.innerHTML = "";
    for (let y = state.today.y + 1; y >= y0; y--) {
      const b = document.createElement("button");
      b.textContent = y;
      b.className = y === state.year ? "on" : "";
      b.onclick = () => { yw.classList.remove("open"); setYear(y); };
      menu.appendChild(b);
    }
    yw.classList.toggle("open");
  };
  document.addEventListener("click", () => yw.classList.remove("open"));
}

init();
