/* ═══════════════════════════════════════════════════════════════
   OVERTIMETAB · веб-версия — логика интерфейса.
   Данные и ПРАВКИ идут через настоящий движок программы (server.py
   импортирует database.py + logic.py). Этап 1: чтение всего главного
   экрана + правка дня (статусы К/Б/О, дежурства, отмена).
   ═══════════════════════════════════════════════════════════════ */

const MONTHS  = ["Январь","Февраль","Март","Апрель","Май","Июнь",
                 "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
const MONTHS_S= ["Янв","Фев","Мар","Апр","Май","Июн","Июл","Авг","Сен","Окт","Ноя","Дек"];
const WEEKDAYS= ["Пн","Вт","Ср","Чт","Пт","Сб","Вс"];
const WD_FULL = ["понедельник","вторник","среда","четверг","пятница","суббота","воскресенье"];

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
  selDate: null,             // открытый в инспекторе день
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

async function post(path, body) {
  const r = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return r.json();
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
   ТОСТЫ — как в программе: тихие капсулы снизу по центру
   ═══════════════════════════════════════════════════════════════ */
function toast(message, kind = "success") {
  const t = document.createElement("div");
  t.className = `toast ${kind === "error" ? "error" : ""}`;
  t.innerHTML = `<span class="dot"></span>${esc(message)}`;
  $("#toasts").appendChild(t);
  setTimeout(() => t.classList.add("out"), kind === "error" ? 3200 : 2200);
  setTimeout(() => t.remove(), kind === "error" ? 3500 : 2500);
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
    if (d.date === state.selDate) cls.push("picked");
    const nCls = ["n"];
    if (d.is_weekend || d.is_holiday) nCls.push("red");
    let flags = "";
    if (d.has_comp) flags += `<span class="badge comp" title="Компенсация (приказ)">В</span>`;
    if (d.status) flags += `<span class="badge st-${d.status}" title="${{К:"Командировка",Б:"Больничный",О:"Отпуск"}[d.status] || d.status}">${d.status}</span>`;
    const duties = d.duties.map(x =>
      `<span class="duty ${x.is_shift ? "shift" : ""}">${esc(x.text)}</span>`).join("");
    const pre = d.is_pre_holiday ? '<span class="pre" title="Предпраздничный день"></span>' : "";
    const [y, m, dd] = d.date.split("-");
    const tip = `${dd}.${m}.${y}` +
      (d.is_holiday ? " · праздник" : d.is_weekend ? " · выходной" : "") +
      (d.is_pre_holiday ? " · предпраздничный" : "") +
      (d.duties.length ? "\nДежурства: " + d.duties.map(x => x.text).join(", ") : "") +
      (d.has_comp ? "\nКомпенсация" : "") + (d.status ? "\nСтатус: " + d.status : "");
    return `<div class="${cls.join(" ")}" data-date="${d.date}" title="${esc(tip)}">
      <div class="top"><span class="${nCls.join(" ")}">${d.n}</span>
        <span class="flags">${flags}</span></div>
      <div class="duties">${duties}</div>${pre}
    </div>`;
  }).join("");
  $("#calendarWrap").className = "calendar-wrap";
  $("#calendarWrap").innerHTML = `
    <div class="cal-head">${wd}</div>
    <div class="cal-grid">${cells}</div>`;
  $("#calendarWrap").querySelectorAll(".day").forEach(el =>
    el.onclick = () => dayClicked(el.dataset.date));
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
   ИНСПЕКТОР ДНЯ — просмотр и первые правки
   ═══════════════════════════════════════════════════════════════ */
const STATUS_INFO = { "К": "Командировка", "Б": "Больничный", "О": "Отпуск" };

function fmtDT(iso) { return iso.slice(11, 16); }
function fmtDM(iso) { return iso.slice(8, 10) + "." + iso.slice(5, 7); }

async function dayClicked(dateIso) {
  if (state.selDate === dateIso) { closeInspector(); return; }
  // клик по «чужому» дню из соседнего месяца — сначала открываем тот месяц
  const m = +dateIso.slice(5, 7);
  if (m !== state.month) { state.month = m; state.view = "month"; await loadMonth(); }
  await openInspector(dateIso);
}

async function openInspector(dateIso) {
  state.selDate = dateIso;
  try {
    const d = await api(`/api/day?emp=${state.empId}&date=${dateIso}`);
    if (state.selDate === dateIso) renderInspector(d);
  } catch (e) {
    toast(e.message, "error");
  }
}

function closeInspector() {
  state.selDate = null;
  $("#inspector").classList.remove("open");
  if (monthCache) renderCalendar(monthCache);
}

function renderInspector(d) {
  const box = $("#inspector");
  const dt = new Date(d.date + "T00:00:00");
  const t = state.today;
  const isToday = d.date === `${t.y}-${String(t.m).padStart(2, "0")}-${String(t.d).padStart(2, "0")}`;
  const pills = [
    d.is_holiday ? '<span class="pill hol">праздник</span>' : "",
    (!d.is_working && !d.is_holiday) ? '<span class="pill">выходной</span>' : "",
    d.is_working && !d.is_holiday ? '<span class="pill">рабочий</span>' : "",
    d.is_pre_holiday ? '<span class="pill pre">предпраздничный</span>' : "",
    isToday ? '<span class="pill tdy">сегодня</span>' : "",
  ].filter(Boolean).join("");

  const stBtn = (code) => `
    <button class="st-btn ${d.status === code ? "on-" + code : ""}" data-st="${code}"
      title="${STATUS_INFO[code]}">
      <span class="k">${code}</span>${STATUS_INFO[code]}</button>`;

  const dutyCards = d.duties.map(x => {
    const cross = x.multi;
    const time = cross
      ? `${fmtDT(x.start)} – ${fmtDT(x.end)}`
      : `${fmtDT(x.start)} – ${fmtDT(x.end)}`;
    const span = cross
      ? `<div class="cm">с ${fmtDM(x.start)} по ${fmtDM(x.end)} · в этот день ${x.slice}</div>` : "";
    const brks = (x.breaks || []).map(b =>
      `<div class="brk">перерыв ${b}</div>`).join("");
    return `
      <div class="duty-card">
        <div class="dr-top">
          <span class="time">${time}</span>
          <span class="chip-s ${x.is_shift ? "" : "ns"}">${x.is_shift ? "сменное" : "сверх нормы"}</span>
          <button class="del" data-id="${x.id}" title="Удалить дежурство">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
              stroke-linecap="round"><path d="M3 6h18M8 6V4h8v2m1 0-1 14H8L7 6"/></svg>
          </button>
        </div>
        ${span}
        ${x.comment ? `<div class="cm">${esc(x.comment)}</div>` : ""}
        ${brks}
      </div>`;
  }).join("");

  const comps = d.comps.map(c =>
    `<div class="brk" style="margin-top:0">${esc(c.text)}</div>`).join("");

  box.innerHTML = `
    <div class="insp-head">
      <div>
        <div class="insp-date">${dt.getDate()} ${MONTHS[dt.getMonth()].toLowerCase()}, ${WD_FULL[dt.getDay() === 0 ? 6 : dt.getDay() - 1]}</div>
        <div class="insp-sub">${pills}</div>
      </div>
      <button class="insp-close" title="Закрыть (Esc)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"
          stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>
      </button>
    </div>
    <div class="insp-body">
      <div class="insp-sec">
        <div class="cap">Статус дня</div>
        <div class="status-row">
          ${stBtn("К")}${stBtn("Б")}${stBtn("О")}
          <button class="st-none" data-st="" title="Снять статус">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"
              stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>
          </button>
        </div>
      </div>

      <div class="insp-sec">
        <div class="cap">Дежурства · ${d.duties.length}</div>
        ${dutyCards || '<div class="cm" style="font-size:11.5px;color:var(--txt3)">В этот день дежурств нет</div>'}
        <button class="add-toggle" id="addToggle">+ Добавить дежурство</button>
        <form class="add-form" id="addForm" style="display:none">
          <div class="f-row">
            <label>Начало</label><input type="time" id="fStart" value="08:00" required>
            <label style="width:auto">Конец</label><input type="time" id="fEnd" value="20:00" required>
          </div>
          <div class="f-row">
            <label>Тип</label>
            <label class="f-check"><input type="checkbox" id="fShift" checked> сменное</label>
            <label style="width:auto">Комментарий</label>
            <input type="text" id="fComment" class="f-txt" placeholder="необязательно" style="flex:1">
          </div>
          <div class="cm" style="font-size:10.5px;color:var(--txt3)">
            Конец раньше начала = дежурство до следующего дня
          </div>
          <div class="form-err" id="formErr"></div>
          <div class="f-row" style="justify-content:flex-end">
            <button type="button" class="btn ghost" id="fCancel">Отмена</button>
            <button type="submit" class="btn primary">Сохранить</button>
          </div>
        </form>
      </div>

      ${comps ? `
      <div class="insp-sec">
        <div class="cap">Компенсации в этот день</div>
        ${comps}
      </div>` : ""}
    </div>
    <div class="insp-foot">
      <button class="undo-btn" id="undoBtn" title="Отменить последнее действие">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
          stroke-linecap="round" stroke-linejoin="round"><path d="M9 14 4 9l5-5"/><path d="M4 9h10a6 6 0 0 1 0 12h-3"/></svg>
        Отменить
      </button>
      <span class="insp-hint">Ctrl+Z тоже работает</span>
    </div>`;
  box.classList.add("open");

  box.querySelector(".insp-close").onclick = closeInspector;

  // статусы
  box.querySelectorAll("[data-st]").forEach(b =>
    b.onclick = () => applyWrite(post("/api/day/set-status",
      { emp: state.empId, date: d.date, status: b.dataset.st })));

  // удаление дежурства: первый клик — подтверждение, второй — удаление
  box.querySelectorAll(".del").forEach(b =>
    b.onclick = () => {
      if (!b.classList.contains("confirm")) {
        b.classList.add("confirm");
        b.textContent = "Точно?";
        setTimeout(() => { if (b.isConnected) {
          b.classList.remove("confirm");
          b.innerHTML = svgTrash();
        }}, 2600);
        return;
      }
      applyWrite(post("/api/duty/delete",
        { id: +b.dataset.id, emp: state.empId, year: state.year, month: state.month }));
    });

  // форма добавления
  const form = box.querySelector("#addForm");
  box.querySelector("#addToggle").onclick = () => {
    form.style.display = form.style.display === "none" ? "flex" : "none";
    box.querySelector("#addToggle").style.display =
      form.style.display === "none" ? "" : "none";
    if (form.style.display !== "none") box.querySelector("#fStart").focus();
  };
  box.querySelector("#fCancel").onclick = () => {
    form.style.display = "none";
    box.querySelector("#addToggle").style.display = "";
    box.querySelector("#formErr").classList.remove("show");
  };
  form.onsubmit = async (e) => {
    e.preventDefault();
    const err = box.querySelector("#formErr");
    err.classList.remove("show");
    const res = await post("/api/duty/add", {
      emp: state.empId,
      date: d.date,
      start: box.querySelector("#fStart").value,
      end: box.querySelector("#fEnd").value,
      is_shift: box.querySelector("#fShift").checked,
      comment: box.querySelector("#fComment").value.trim(),
    });
    if (!res.ok) {
      err.textContent = res.message || "Ошибка сохранения";
      err.classList.add("show");
      return;
    }
    afterWrite(res);
  };

  box.querySelector("#undoBtn").onclick = doUndo;
}

function svgTrash() {
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"
    stroke-linecap="round"><path d="M3 6h18M8 6V4h8v2m1 0-1 14H8L7 6"/></svg>`;
}

/* Единая обработка ответов правок: тост + свежий месяц + инспектор */
function afterWrite(res) {
  toast(res.message, res.ok ? "success" : "error");
  if (!res.ok) return;
  if (res.reload) { hardReload(); return; }
  if (res.month) {
    applyMonth(res.month);
    if (state.selDate) openInspector(state.selDate);
  }
}

async function applyWrite(promise) {
  try { afterWrite(await promise); }
  catch (e) { toast("Ошибка: " + e.message, "error"); }
}

async function doUndo() {
  applyWrite(post("/api/undo", {}));
}

async function hardReload() {
  await loadBootstrap();
  await loadMonth();
  if (state.selDate) openInspector(state.selDate);
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

function applyMonth(d) {
  monthCache = d;
  renderCalendar(d);
  renderSummary(d);
  const i = state.employees.findIndex(e => e.id === d.emp.id);
  if (i >= 0) {
    state.employees[i].ratio = d.ratio;
    state.employees[i].mini = d.mini;
  }
  recalcMiniMax();
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
    applyMonth(d);
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
  closeInspector();
  renderList();
  state.view === "year" ? loadYear() : loadMonth();
}

function gotoMonth(m) {
  if (state.month === m && state.view === "month") return;
  state.month = m;
  state.view = "month";
  closeInspector();
  loadMonth();
}

function showYear() {
  state.view = "year";
  closeInspector();
  loadYear();
}

async function setYear(y) {
  if (state.year === y) return;
  state.year = y;
  $("#yearWrap")?.classList.remove("open");
  closeInspector();
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
        <div><b style="color:#2D3B45">Нет связи с сервером веб-версии.</b><br><br>
        Откройте живое превью процесса «Веб-версия OVERTIMETAB» (порт 8081) —<br>
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
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeInspector();
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
      e.preventDefault();
      doUndo();
    }
  });

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
