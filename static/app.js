const $ = (id) => document.getElementById(id);

let DATA = null;
let SETTINGS = null;
let RESULT = null;
let DATA_SOURCE = null;

let currentBlockClass = "";
let currentFixedClass = "";
let currentResultClass = "";
let currentTeacher = "";
let currentSubject = "";
let dragSource = null;
let saveTimer = null;
let pendingTab = null;
let solveVariant = 0;
let pendingDrag = null;

const esc = (value) =>
  String(value ?? "").replace(/[&<>"']/g, (ch) => {
    const map = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
    return map[ch];
  });

const slotKey = (day, period) => `${day}-${period}`;
const parseSlot = (key) => key.split("-").map(Number);

async function apiGet(url) {
  const response = await fetch(url);
  return response.json();
}

async function apiPost(url, body) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return response.json();
}

function gradeOf(cls) {
  const match = cls.match(/^[\u4e00-\u9fa5]+?(?=[0-9一二三四五六七八九十]|$)/);
  if (match && match[0]) return match[0];
  return cls.slice(0, 1);
}

function sameGradeClasses(cls) {
  const grade = gradeOf(cls);
  return DATA.classes.filter((item) => gradeOf(item) === grade);
}

function slotText(key) {
  const [day, period] = parseSlot(key);
  return `${DATA.days[day]}第${period + 1}节`;
}

function sectionCellHtml(periodIndex) {
  const span = DATA.sectionSpans ? DATA.sectionSpans[periodIndex] : 0;
  if (!span) return "";
  return `<td rowspan="${span}">${esc(DATA.periods[periodIndex].section || "")}</td>`;
}

function isBlocked(cls, key) {
  return (SETTINGS.classBlocked[cls] || []).includes(key);
}

function hasFixed(cls, key) {
  return SETTINGS.fixed[cls] && SETTINGS.fixed[cls][key];
}

function hasManual(cls, key) {
  return SETTINGS.manual[cls] && SETTINGS.manual[cls][key];
}

function teacherBlocked(teacher, key) {
  return (SETTINGS.teacherUnavailable[teacher] || []).includes(key);
}

function blockCountFor(cls) {
  return (SETTINGS.classBlocked[cls] || []).length;
}

function blockTargetFor(cls) {
  const totalHours = DATA.subjects.reduce((sum, subject) => sum + (DATA.hours[cls][subject] || 0), 0);
  return Math.max(0, DATA.availableSlots - totalHours);
}

function insufficientBlockClasses() {
  return DATA.classes.filter((cls) => blockCountFor(cls) < blockTargetFor(cls));
}

let toastTimer = null;
function showToast(message) {
  const toast = $("toast");
  toast.textContent = message;
  toast.classList.remove("hidden");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.add("hidden"), 2000);
}

const BACKUP_KEY = "paike_backup_v1";
const AUTO_UPLOAD_KEY = "paike_auto_uploaded";

function readBackup() {
  try {
    return JSON.parse(localStorage.getItem(BACKUP_KEY) || "null") || {};
  } catch (error) {
    return {};
  }
}

function writeBackup(partial) {
  const current = readBackup();
  const next = { ...current, ...partial };
  try {
    localStorage.setItem(BACKUP_KEY, JSON.stringify(next));
  } catch (error) {
    console.warn("浏览器备份保存失败", error);
  }
}

function clearBackup() {
  localStorage.removeItem(BACKUP_KEY);
  sessionStorage.removeItem(AUTO_UPLOAD_KEY);
}

function mergeSavedSettings(saved, current) {
  const result = {
    ...current,
    ...saved,
    rules: { ...current.rules, ...(saved.rules || {}) },
    titles: { ...current.titles, ...(saved.titles || {}) },
    preferredConsecutive: {
      ...current.preferredConsecutive,
      ...(saved.preferredConsecutive || {}),
    },
  };
  return result;
}

async function init() {
  const payload = await apiGet("/api/data");
  if (!payload.ok) {
    $("dataStatus").textContent = `读取失败：${payload.error || "未知错误"}`;
    return;
  }
  DATA = payload.data;
  SETTINGS = payload.settings;
  DATA_SOURCE = payload.dataSource || { name: "基本数据.xlsx", isDemo: true };

  const backup = readBackup();
  if (backup.contentBase64 && !backup.data) {
    const parsePayload = await apiPost("/api/parse", {
      filename: backup.filename || "upload.xlsx",
      contentBase64: backup.contentBase64,
    });
    if (parsePayload.ok) {
      backup.data = parsePayload.data;
      writeBackup({ data: parsePayload.data });
    }
  }
  if (backup.contentBase64 && backup.data) {
    DATA = backup.data;
    DATA_SOURCE = { name: backup.filename || "上传数据.xlsx", isDemo: false };
    SETTINGS = mergeSavedSettings(backup.settings || {}, SETTINGS);
  } else if (backup.settings) {
    SETTINGS = mergeSavedSettings(backup.settings, SETTINGS);
  }
  sessionStorage.removeItem(AUTO_UPLOAD_KEY);

  $("dataStatus").textContent = "已读取";
  populateClassSelect("blockClassSelect", setBlockClass);
  populateClassSelect("fixedClassSelect", setFixedClass);
  populateClassSelect("resultClassSelect", setResultClass);
  populateSubjectSelect();
  populateTeacherSelect();
  renderOverview();
  renderDataSourceInfo();
  renderBlockGrid();
  renderFixedGrid();
  renderSubjectGrid();
  renderTeacherGrid();
  renderRules();
  renderTitleInputs();
  try {
    const usagePayload = await apiGet("/api/usage");
    renderUsage(usagePayload.usage || {});
  } catch (error) {
    console.error(error);
  }
  attachEvents();
}

function renderUsage(usage) {
  for (const [tab, lines] of Object.entries(usage)) {
    const container = $(`usage-${tab}`);
    if (!container) continue;
    container.innerHTML = (lines || []).map((line) => `<div>${esc(line)}</div>`).join("");
  }
}

function populateClassSelect(id, onChange) {
  const select = $(id);
  select.innerHTML = DATA.classes
    .map((cls) => `<option value="${esc(cls)}">${esc(cls)}</option>`)
    .join("");
  select.addEventListener("change", () => onChange(select.value));
  onChange(select.value);
}

function populateTeacherSelect() {
  currentTeacher = DATA.teachers[0] || "";
  const options = DATA.teachers
    .map((teacher) => `<option value="${esc(teacher)}">${esc(teacher)}</option>`)
    .join("");
  $("teacherSelect").innerHTML = options;
  $("resultTeacherSelect").innerHTML = options;
  $("teacherSelect").value = currentTeacher;
  $("resultTeacherSelect").value = currentTeacher;
  $("teacherSelect").addEventListener("change", () => {
    currentTeacher = $("teacherSelect").value;
    renderTeacherGrid();
  });
  $("resultTeacherSelect").addEventListener("change", () => {
    currentTeacher = $("resultTeacherSelect").value;
    renderTeacherResultGrid();
  });
}

function populateSubjectSelect() {
  currentSubject = DATA.subjects[0] || "";
  $("subjectSelect").innerHTML = DATA.subjects
    .map((subject) => `<option value="${esc(subject)}">${esc(subject)}</option>`)
    .join("");
  $("subjectSelect").value = currentSubject;
  $("subjectSelect").addEventListener("change", () => {
    currentSubject = $("subjectSelect").value;
    renderSubjectGrid();
  });
}

function renderOverview() {
  const totals = DATA.classes.map((cls) =>
    DATA.subjects.reduce((sum, subject) => sum + (DATA.hours[cls][subject] || 0), 0)
  );
  const cardData = [
    ["班级数", DATA.classes.length],
    ["科目数", DATA.subjects.length],
    ["教师数", DATA.teachers.length],
    ["每周可用格", DATA.availableSlots],
    ["每班周课时", totals[0] || 0],
  ];
  $("summaryCards").innerHTML = cardData
    .map(
      ([label, num]) => `
        <div class="stat-card">
          <div class="label">${esc(label)}</div>
          <div class="num">${esc(num)}</div>
        </div>`
    )
    .join("");

  let html = "<thead><tr><th>时段</th><th>节次</th><th>时间</th>";
  DATA.days.forEach((day) => (html += `<th>${esc(day)}</th>`));
  html += "</tr></thead><tbody>";
  DATA.periods.forEach((period, periodIndex) => {
    html += `<tr>${sectionCellHtml(periodIndex)}<td>${esc(period.period)}</td><td>${esc(period.time)}</td>`;
    DATA.days.forEach(() => (html += "<td></td>"));
    html += "</tr>";
  });
  html += "</tbody>";
  $("scheduleTable").innerHTML = html;

  $("dataWarnings").innerHTML = (DATA.warnings || [])
    .map((warning) => `<div>${esc(warning)}</div>`)
    .join("");
}

function renderDataSourceInfo() {
  const info = DATA_SOURCE || {};
  if (info.isDemo) {
    $("dataSourceInfo").innerHTML = "当前为演示数据，可上传自己的表格后立即使用。";
  } else {
    $("dataSourceInfo").innerHTML = `已读取 ${esc(info.name)}`;
  }
}

function renderTitleInputs() {
  $("classTitleInput").value = SETTINGS.titles.classSuffix || "";
  $("teacherTitleInput").value = SETTINGS.titles.teacherSuffix || "";
}

function renderBlockGrid() {
  const cls = currentBlockClass;
  if (!cls) return;
  let html = "<table class='schedule-grid'><thead><tr><th>时段</th><th>节次</th><th>时间</th>";
  DATA.days.forEach((day) => (html += `<th>${esc(day)}</th>`));
  html += "</tr></thead><tbody>";

  DATA.periods.forEach((period, periodIndex) => {
    html += `<tr>${sectionCellHtml(periodIndex)}<td>${esc(period.period)}</td><td>${esc(period.time)}</td>`;
    DATA.days.forEach((_, dayIndex) => {
      const key = slotKey(dayIndex, periodIndex);
      const blocked = isBlocked(cls, key);
      html += `<td class="cell clickable ${blocked ? "blocked" : "empty"}" data-key="${key}">${
        blocked ? "不排课" : ""
      }</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  $("blockGrid").innerHTML = html;
  const count = (SETTINGS.classBlocked[cls] || []).length;
  const totalHours = DATA.subjects.reduce((sum, subject) => sum + (DATA.hours[cls][subject] || 0), 0);
  const totalSlots = DATA.availableSlots;
  const target = Math.max(0, totalSlots - totalHours);
  $("blockCount").innerHTML =
    `<div>${esc(cls)}周课程量为${totalHours}，总空格数为${totalSlots}，请设置${target}个空格为不排课</div>` +
    `<div>${esc(cls)}已设置${count}个不排课</div>`;
}

function setBlockClass(cls) {
  currentBlockClass = cls;
  renderBlockGrid();
}

function renderFixedGrid() {
  const cls = currentFixedClass;
  if (!cls) return;
  let html = "<table class='schedule-grid'><thead><tr><th>时段</th><th>节次</th><th>时间</th>";
  DATA.days.forEach((day) => (html += `<th>${esc(day)}</th>`));
  html += "</tr></thead><tbody>";

  DATA.periods.forEach((period, periodIndex) => {
    html += `<tr>${sectionCellHtml(periodIndex)}<td>${esc(period.period)}</td><td>${esc(period.time)}</td>`;
    DATA.days.forEach((_, dayIndex) => {
      const key = slotKey(dayIndex, periodIndex);
      const fixedSubject = hasFixed(cls, key);
      const manualSubject = hasManual(cls, key);
      const subject = fixedSubject || manualSubject;
      const teacher = subject ? DATA.teacherMap[cls][subject] || "" : "";
      const blocked = isBlocked(cls, key);
      let cellClass = "cell clickable";
      let text = "";
      if (blocked) {
        cellClass += " blocked";
        text = "不排课";
      } else if (subject) {
        cellClass += fixedSubject ? " fixed" : " manual";
        text = `<span class="subject">${esc(subject)}</span><span class="teacher">${esc(teacher)}</span>`;
      } else {
        cellClass += " empty";
      }
      html += `<td class="${cellClass}" data-key="${key}" data-blocked="${blocked ? "1" : ""}">${text}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  $("fixedGrid").innerHTML = html;

  const count =
    Object.keys(SETTINGS.fixed[cls] || {}).length +
    Object.keys(SETTINGS.manual[cls] || {}).length;
  $("fixedCount").textContent = `${cls}已设置${count}个固定时间段课程`;
}

function openFixedModal(key) {
  const cls = currentFixedClass;
  if (!cls) return;
  const [day, period] = parseSlot(key);
  $("fixedModalTitle").textContent = `${cls} ${DATA.days[day]}第${period + 1}节`;
  $("fixedModal").dataset.key = key;
  const subjects = DATA.subjects.filter(
    (subject) => DATA.hours[cls][subject] > 0 && DATA.teacherMap[cls][subject]
  );
  $("fixedModalList").innerHTML = subjects
    .map(
      (subject) => `
        <div class="course-option" data-subject="${esc(subject)}">
          <span class="course-name">${esc(subject)}</span>
          <span class="scopes">
            <label><input type="checkbox" data-scope="grade"> 全年级</label>
            <label><input type="checkbox" data-scope="all"> 全校</label>
          </span>
        </div>`
    )
    .join("");
  $("fixedModal").classList.add("show");
}

function closeFixedModal() {
  $("fixedModal").classList.remove("show");
}

function applyFixedModalChoice(subject, scope) {
  const cls = currentFixedClass;
  const key = $("fixedModal").dataset.key;
  if (!cls || !key || !subject) return;
  let targets = [cls];
  if (scope === "grade") targets = sameGradeClasses(cls);
  if (scope === "all") targets = DATA.classes;
  let applied = 0;
  targets.forEach((target) => {
    if (!DATA.hours[target][subject] || !DATA.teacherMap[target][subject]) return;
    toggleFixed(target, key, subject);
    applied += 1;
  });
  closeFixedModal();
  renderFixedGrid();
  saveSettings();
  return applied;
}

function setFixedClass(cls) {
  currentFixedClass = cls;
  renderFixedGrid();
}

function setResultClass(cls) {
  currentResultClass = cls;
  renderResultClassGrid();
}

function renderSubjectGrid() {
  const subject = currentSubject;
  if (!subject) return;
  let html = "<table class='schedule-grid'><thead><tr><th>时段</th><th>节次</th><th>时间</th>";
  DATA.days.forEach((day) => (html += `<th>${esc(day)}</th>`));
  html += "</tr></thead><tbody>";

  DATA.periods.forEach((period, periodIndex) => {
    html += `<tr>${sectionCellHtml(periodIndex)}<td>${esc(period.period)}</td><td>${esc(period.time)}</td>`;
    DATA.days.forEach((_, dayIndex) => {
      const key = slotKey(dayIndex, periodIndex);
      const blocked = (SETTINGS.subjectConstraints[subject] || []).includes(key);
      let cellClass = "cell clickable empty";
      let text = "";
      if (blocked) {
        cellClass = "cell clickable teacher-off";
        text = "不排课";
      }
      html += `<td class="${cellClass}" data-key="${key}">${text}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  $("subjectGrid").innerHTML = html;
  const count = (SETTINGS.subjectConstraints[subject] || []).length;
  $("subjectCount").textContent = `${subject} 已设置 ${count} 个不出现时段`;
}

function renderTeacherGrid() {
  const teacher = currentTeacher;
  if (!teacher) return;
  let html = "<table class='schedule-grid'><thead><tr><th>时段</th><th>节次</th><th>时间</th>";
  DATA.days.forEach((day) => (html += `<th>${esc(day)}</th>`));
  html += "</tr></thead><tbody>";

  DATA.periods.forEach((period, periodIndex) => {
    html += `<tr>${sectionCellHtml(periodIndex)}<td>${esc(period.period)}</td><td>${esc(period.time)}</td>`;
    DATA.days.forEach((_, dayIndex) => {
      const key = slotKey(dayIndex, periodIndex);
      const blocked = teacherBlocked(teacher, key);
      let cellClass = "cell clickable empty";
      let text = "";
      if (blocked) {
        cellClass = "cell clickable teacher-off";
        text = "不排课";
      }
      html += `<td class="${cellClass}" data-key="${key}">${text}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  $("teacherGrid").innerHTML = html;
  const count = (SETTINGS.teacherUnavailable[teacher] || []).length;
  $("teacherSelectedCount").textContent = `${teacher} 已设置 ${count} 个不排课时间格`;
}

function renderRules() {
  $("dailyMaxInput").value = SETTINGS.rules.dailyMax;
  $("consecutiveMaxInput").value = SETTINGS.rules.consecutiveMax;
  $("preferredConsecutiveInput").value = 2;
  renderPreferredSubjectList();
}

function renderPreferredSubjectList() {
  $("preferredSubjectList").innerHTML = DATA.subjects
    .map((subject) => {
      const count = SETTINGS.preferredConsecutive[subject] || 0;
      return `<label>
        <input type="checkbox" value="${esc(subject)}" ${count ? "checked" : ""}>
        <span>${esc(subject)}</span>
        ${count ? `<span class="load">${count}节</span>` : ""}
      </label>`;
    })
    .join("");
}

function applyPreferredSelection() {
  const requested = Math.max(1, Math.min(12, Number($("preferredConsecutiveInput").value) || 2));
  const max = SETTINGS.rules.consecutiveMax;
  const count = Math.min(requested, max);
  const selected = [...document.querySelectorAll("#preferredSubjectList input:checked")].map((input) => input.value);
  SETTINGS.preferredConsecutive = {};
  selected.forEach((subject) => {
    SETTINGS.preferredConsecutive[subject] = count;
  });
  renderPreferredSubjectList();
}

function renderResult() {
  if (!RESULT) return;
  $("exportButton").disabled = false;
  $("resetResultButton").disabled = false;
  if (!currentResultClass) {
    currentResultClass = DATA.classes[0];
    $("resultClassSelect").value = currentResultClass;
  }
  if (!currentTeacher) {
    currentTeacher = DATA.teachers[0];
    $("resultTeacherSelect").value = currentTeacher;
  }
  renderResultClassGrid();
  renderTeacherResultGrid();
  renderReport();
  if (RESULT.status === "OPTIMAL") {
    $("solveStatus").textContent = "排课完成：OPTIMAL，可在班级课表中拖动调课。";
  } else {
    $("solveStatus").textContent = "排课完成：FEASIBLE，可在班级课表中拖动调课。";
  }
}

function renderResultClassGrid() {
  const cls = currentResultClass;
  if (!cls || !RESULT) return;
  let html = "<table class='schedule-grid'><thead><tr><th>时段</th><th>节次</th><th>时间</th>";
  DATA.days.forEach((day) => (html += `<th>${esc(day)}</th>`));
  html += "</tr></thead><tbody>";

  DATA.periods.forEach((period, periodIndex) => {
    html += `<tr>${sectionCellHtml(periodIndex)}<td>${esc(period.period)}</td><td>${esc(period.time)}</td>`;
    DATA.days.forEach((_, dayIndex) => {
      const key = slotKey(dayIndex, periodIndex);
      const cell = RESULT.schedule[cls] ? RESULT.schedule[cls][key] : null;
      const blocked = isBlocked(cls, key);
      let cellClass = "cell";
      let text = "";
      let draggable = "";
      let dataSubject = "";
      let dataSource = "";
      let dataBlocked = "";
      if (blocked) {
        cellClass += " blocked";
        text = "不排课";
        dataBlocked = "1";
      } else if (cell) {
        dataSubject = cell.subject;
        dataSource = cell.source;
        cellClass += ` ${cell.source}`;
        if (cell.source !== "fixed") {
          draggable = "draggable='true'";
        }
        const badge = cell.source === "fixed" ? "<span class='badge'>固定</span>" : cell.source === "manual" ? "<span class='badge'>手动</span>" : "";
        text = `<span class="subject">${esc(cell.subject)}</span><span class="teacher">${esc(cell.teacher)}</span>${badge}`;
      } else {
        cellClass += " empty";
      }
      html += `<td class="${cellClass}" data-key="${key}" data-subject="${esc(dataSubject)}" data-source="${esc(
        dataSource
      )}" data-blocked="${dataBlocked}" ${draggable}>${text}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  $("resultClassGrid").innerHTML = html;
}

function renderTeacherResultGrid() {
  const teacher = currentTeacher;
  if (!teacher || !RESULT) return;
  let html = "<table class='schedule-grid'><thead><tr><th>时段</th><th>节次</th><th>时间</th>";
  DATA.days.forEach((day) => (html += `<th>${esc(day)}</th>`));
  html += "</tr></thead><tbody>";

  DATA.periods.forEach((period, periodIndex) => {
    html += `<tr>${sectionCellHtml(periodIndex)}<td>${esc(period.period)}</td><td>${esc(period.time)}</td>`;
    DATA.days.forEach((_, dayIndex) => {
      const key = slotKey(dayIndex, periodIndex);
      const entries = [];
      for (const cls of DATA.classes) {
        const cell = RESULT.schedule[cls] ? RESULT.schedule[cls][key] : null;
        if (cell && cell.teacher === teacher) {
          entries.push(`${cls}·${cell.subject}`);
        }
      }
      if (teacherBlocked(teacher, key)) {
        html += `<td class="cell teacher-off">不排课</td>`;
      } else if (entries.length) {
        html += `<td class="cell fixed">${entries.map(esc).join("<br>")}</td>`;
      } else {
        html += `<td class="cell empty"></td>`;
      }
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  $("resultTeacherGrid").innerHTML = html;
}

function renderReport() {
  if (!RESULT) return;
  const issues = validateResultLocal();
  const reportLines = (RESULT.report || []).map((line) => `<div>${esc(line)}</div>`).join("");
  const issueLines = issues.length
    ? `<div class="section-title">发现 ${issues.length} 个问题</div>` + issues.map((issue) => `<div class="issue">${esc(issue)}</div>`).join("")
    : `<div style="color:var(--green)">未发现问题</div>`;
  $("resultReport").innerHTML = reportLines + issueLines;
}

function validateResultLocal(schedule) {
  const result = schedule || (RESULT ? RESULT.schedule : {});
  const issues = [];
  const counts = {};
  for (const cls of DATA.classes) {
    counts[cls] = {};
    DATA.subjects.forEach((subject) => (counts[cls][subject] = 0));
    const cells = result[cls] || {};
    for (const [key, cell] of Object.entries(cells)) {
      const subject = cell.subject;
      counts[cls][subject] = (counts[cls][subject] || 0) + 1;
      if (isBlocked(cls, key)) issues.push(`${cls} ${slotText(key)} 标为不排课，但仍有课程`);
      if (cell.teacher && teacherBlocked(cell.teacher, key)) {
        issues.push(`教师 ${cell.teacher} 在 ${slotText(key)} 不排课，但 ${cls} 仍安排了课程`);
      }
      if ((SETTINGS.subjectConstraints[subject] || []).includes(key)) {
        issues.push(`${cls} 的 ${subject} 出现在科目时段约束 ${slotText(key)}`);
      }
      const expectedTeacher = DATA.teacherMap[cls][subject];
      if (expectedTeacher && cell.teacher !== expectedTeacher) {
        issues.push(`${cls} ${slotText(key)} 的 ${subject} 教师与任课表不一致`);
      }
    }
  }
  for (const cls of DATA.classes) {
    for (const subject of DATA.subjects) {
      const actual = counts[cls][subject] || 0;
      const expected = DATA.hours[cls][subject];
      if (actual !== expected) issues.push(`${cls} 的 ${subject} 实际 ${actual} 节，应排 ${expected} 节`);
    }
  }

  const occupied = {};
  for (const cls of DATA.classes) {
    for (const [key, cell] of Object.entries(result[cls] || {})) {
      if (!cell.teacher) continue;
      const id = `${cell.teacher}|${key}`;
      if (occupied[id]) {
        issues.push(`教师 ${cell.teacher} 在 ${slotText(key)} 同时出现在 ${occupied[id]} 和 ${cls}`);
      } else {
        occupied[id] = cls;
      }
    }
  }

  const dailyMax = SETTINGS.rules.dailyMax;
  const consecutiveMax = SETTINGS.rules.consecutiveMax;
  for (const cls of DATA.classes) {
    const classCells = result[cls] || {};
    for (let day = 0; day < DATA.days.length; day++) {
      for (const subject of DATA.subjects) {
        let dayCount = 0;
        for (let period = 0; period < DATA.nSlots; period++) {
          const cell = classCells[slotKey(day, period)];
          if (cell && cell.subject === subject) dayCount += 1;
        }
        if (dayCount > dailyMax) {
          issues.push(`${cls} ${DATA.days[day]} 的 ${subject} 达到 ${dayCount} 节，超过上限 ${dailyMax}`);
        }
      }
      let runCells = [];
      const checkRun = () => {
        if (!runCells.length) return;
        if (
          runCells.length > consecutiveMax &&
          runCells.some((cell) => cell.source !== "fixed")
        ) {
          issues.push(
            `${cls} ${DATA.days[day]} 的 ${runCells[0].subject} 连堂达到 ${runCells.length} 节，超过上限 ${consecutiveMax}`
          );
        }
      };
      for (let period = 0; period < DATA.nSlots; period++) {
        const cell = classCells[slotKey(day, period)];
        if (cell && runCells.length && runCells[runCells.length - 1].subject === cell.subject) {
          runCells.push(cell);
        } else {
          checkRun();
          runCells = cell ? [cell] : [];
        }
      }
      checkRun();
    }
  }
  return [...new Set(issues)];
}

function toggleBlock(cls, key) {
  const list = SETTINGS.classBlocked[cls] || (SETTINGS.classBlocked[cls] = []);
  const index = list.indexOf(key);
  if (index >= 0) {
    list.splice(index, 1);
  } else {
    list.push(key);
    if (SETTINGS.fixed[cls]) delete SETTINGS.fixed[cls][key];
    if (SETTINGS.manual[cls]) delete SETTINGS.manual[cls][key];
  }
}

function copyBlockPattern(source, targets) {
  const pattern = [...(SETTINGS.classBlocked[source] || [])];
  targets.forEach((cls) => {
    SETTINGS.classBlocked[cls] = [...pattern];
  });
}

function toggleFixed(cls, key, subject) {
  if (!SETTINGS.fixed[cls]) SETTINGS.fixed[cls] = {};
  if (subject === "__clear__") {
    delete SETTINGS.fixed[cls][key];
    if (SETTINGS.manual[cls]) delete SETTINGS.manual[cls][key];
  } else {
    SETTINGS.fixed[cls][key] = subject;
    const blocked = SETTINGS.classBlocked[cls] || [];
    const index = blocked.indexOf(key);
    if (index >= 0) blocked.splice(index, 1);
    if (SETTINGS.manual[cls]) delete SETTINGS.manual[cls][key];
  }
}

function deleteManual(cls, key) {
  if (!SETTINGS.manual[cls]) return;
  delete SETTINGS.manual[cls][key];
  if (Object.keys(SETTINGS.manual[cls]).length === 0) delete SETTINGS.manual[cls];
}

function saveSettings() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    writeBackup({ settings: SETTINGS });
  }, 250);
}

function showErrors(errors) {
  const box = $("solveErrors");
  if (errors && errors.length) {
    box.innerHTML = errors.map((item) => `<div>${esc(item)}</div>`).join("");
    box.classList.add("show");
  } else {
    box.innerHTML = "";
    box.classList.remove("show");
  }
}

async function runSolve() {
  if (!DATA) return;
  $("solveStatus").textContent = "正在排课...最长3分钟...";
  showErrors([]);
  $("solveButton").disabled = true;
  try {
    solveVariant += 1;
    const backup = readBackup();
    const solveBody = { settings: SETTINGS, variant: solveVariant };
    if (backup.contentBase64) {
      solveBody.fileBase64 = backup.contentBase64;
      solveBody.filename = backup.filename || "upload.xlsx";
    }
    const payload = await apiPost("/api/solve", solveBody);
    if (!payload.ok) {
      if (payload.errorKind === "timeout") {
        $("solveStatus").textContent = "排课超时：未在 3 分钟内找到方案，请稍后重试。";
      } else if (payload.errorKind === "infeasible") {
        $("solveStatus").textContent = "排课失败：当前约束下没有可行课表。";
      } else {
        $("solveStatus").textContent = "排课失败：请先检查设置。";
      }
      showErrors(payload.errors || payload.report || []);
      return;
    }
    const missingClasses = DATA.classes.filter(
      (cls) => !payload.schedule || !payload.schedule[cls]
    );
    if (missingClasses.length) {
      $("solveStatus").textContent = "浏览器备份与服务器解析结果不一致，请重新选择文件。";
      showErrors([`以下班级在服务器结果中不存在：${missingClasses.join("、")}`]);
      return;
    }
    RESULT = payload;
    renderResult();
    saveSettings();
  } catch (error) {
    $("solveStatus").textContent = "排课失败";
    showErrors([String(error)]);
  } finally {
    $("solveButton").disabled = false;
  }
}

function resetResult() {
  RESULT = null;
  SETTINGS.manual = {};
  solveVariant += 1;
  $("resultClassGrid").innerHTML = "";
  $("resultTeacherGrid").innerHTML = "";
  $("resultReport").innerHTML = "";
  $("exportButton").disabled = true;
  $("resetResultButton").disabled = true;
  $("solveStatus").textContent = "";
  showErrors([]);
}

async function exportWorkbook() {
  if (!RESULT) return;
  SETTINGS.titles.classSuffix = $("classTitleInput").value.trim();
  SETTINGS.titles.teacherSuffix = $("teacherTitleInput").value.trim();
  const backup = readBackup();
  const exportBody = {
    settings: SETTINGS,
    result: RESULT.schedule,
  };
  if (backup.contentBase64) {
    exportBody.fileBase64 = backup.contentBase64;
    exportBody.filename = backup.filename || "upload.xlsx";
  }
  const payload = await apiPost("/api/export", exportBody);
  if (!payload.ok) {
    showErrors([payload.error || "导出失败"]);
    return;
  }
  for (const file of payload.files || []) {
    const filename = file.path.split(/[\\/]/).pop();
    const response = await fetch("/api/download", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filename }),
    });
    if (!response.ok) {
      showErrors([`下载失败：${filename}`]);
      return;
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    await new Promise((resolve) => setTimeout(resolve, 350));
  }
  $("solveStatus").textContent =
    payload.issues && payload.issues.length
      ? `导出成功，但有 ${payload.issues.length} 个校验问题`
      : "导出成功";
}

function moveOnSchedule(schedule, cls, sourceKey, targetKey) {
  const sourceCell = schedule[cls][sourceKey];
  const targetCell = schedule[cls][targetKey];
  if (!sourceCell || sourceCell.source === "fixed") return false;
  if (targetCell && targetCell.source === "fixed") return false;
  const newSource = targetCell ? { ...targetCell, source: "manual" } : null;
  const newTarget = { ...sourceCell, source: "manual" };
  if (newSource) schedule[cls][sourceKey] = newSource;
  else delete schedule[cls][sourceKey];
  schedule[cls][targetKey] = newTarget;
  return true;
}

function isHardIssue(text) {
  return (
    /同时出现/.test(text) ||
    /标为不排课，但仍有课程/.test(text) ||
    /与任课表不一致/.test(text) ||
    /实际 .* 节，应排/.test(text) ||
    /与锁定课程.*不一致/.test(text)
  );
}

function classifyDragIssues(issues) {
  const hard = [];
  const soft = [];
  issues.forEach((issue) => {
    if (isHardIssue(issue)) hard.push(issue);
    else soft.push(issue);
  });
  return { hard, soft };
}

function applyDragMove(sourceTd, targetTd) {
  const cls = currentResultClass;
  const sourceKey = sourceTd.dataset.key;
  const targetKey = targetTd.dataset.key;
  if (!sourceKey || !targetKey || sourceKey === targetKey) return;
  deleteManual(cls, sourceKey);
  deleteManual(cls, targetKey);
  const moved = moveOnSchedule(RESULT.schedule, cls, sourceKey, targetKey);
  if (!moved) return;
  const sourceSubject = RESULT.schedule[cls][targetKey].subject;
  SETTINGS.manual[cls] = SETTINGS.manual[cls] || {};
  SETTINGS.manual[cls][targetKey] = sourceSubject;
  if (RESULT.schedule[cls][sourceKey]) {
    SETTINGS.manual[cls][sourceKey] = RESULT.schedule[cls][sourceKey].subject;
  }
}

function startDragMove(sourceTd, targetTd) {
  const cls = currentResultClass;
  const sourceKey = sourceTd.dataset.key;
  const targetKey = targetTd.dataset.key;
  if (!sourceKey || !targetKey || sourceKey === targetKey) return;
  if (targetTd.dataset.blocked === "1") {
    showToast("不能放入不排课时间格");
    return;
  }
  const trial = JSON.parse(JSON.stringify(RESULT.schedule));
  if (!moveOnSchedule(trial, cls, sourceKey, targetKey)) {
    showToast("不能移动固定课，也不能覆盖固定课");
    return;
  }
  const issues = validateResultLocal(trial);
  const { hard, soft } = classifyDragIssues(issues);
  if (hard.length) {
    showToast(`调整被禁止：${hard[0]}`);
    return;
  }
  if (soft.length) {
    pendingDrag = { sourceTd, targetTd };
    $("dragConfirmText").innerHTML =
      `<div>本次调整会触发以下提示，确认执行吗？</div>` +
      soft.map((issue) => `<div class="issue">${esc(issue)}</div>`).join("");
    $("dragConfirmModal").classList.add("show");
    return;
  }
  applyDragMove(sourceTd, targetTd);
  renderResultClassGrid();
  renderReport();
  saveSettings();
}

function attachEvents() {
  document.querySelectorAll("#tabs .tab").forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });
  document.querySelectorAll("#resultTabs .result-tab").forEach((button) => {
    button.addEventListener("click", () => switchResultTab(button.dataset.view));
  });

  $("blockGrid").addEventListener("click", (event) => {
    const td = event.target.closest("td[data-key]");
    if (!td || !currentBlockClass) return;
    const key = td.dataset.key;
    const adding = !isBlocked(currentBlockClass, key);
    if (adding && blockCountFor(currentBlockClass) >= blockTargetFor(currentBlockClass)) {
      showToast(`该班不排课量应为${blockTargetFor(currentBlockClass)}`);
      return;
    }
    toggleBlock(currentBlockClass, key);
    if ($("blockApplyAll").checked) {
      copyBlockPattern(currentBlockClass, DATA.classes);
      showToast(`已同步到全部 ${DATA.classes.length} 个班`);
    } else if ($("blockApplyGrade").checked) {
      const targets = sameGradeClasses(currentBlockClass);
      copyBlockPattern(currentBlockClass, targets);
      showToast(`已同步到同年级 ${targets.length} 个班`);
    }
    renderBlockGrid();
    saveSettings();
  });

  $("fixedGrid").addEventListener("click", (event) => {
    const td = event.target.closest("td[data-key]");
    if (!td || !currentFixedClass) return;
    if (td.dataset.blocked === "1") return;
    openFixedModal(td.dataset.key);
  });

  $("clearBlockClass").addEventListener("click", () => {
    if (!currentBlockClass) return;
    delete SETTINGS.classBlocked[currentBlockClass];
    renderBlockGrid();
    saveSettings();
  });

  $("clearAllBlock").addEventListener("click", () => {
    SETTINGS.classBlocked = {};
    renderBlockGrid();
    saveSettings();
  });

  $("clearFixedClass").addEventListener("click", () => {
    if (!currentFixedClass) return;
    delete SETTINGS.fixed[currentFixedClass];
    delete SETTINGS.manual[currentFixedClass];
    renderFixedGrid();
    saveSettings();
  });

  $("clearAllFixed").addEventListener("click", () => {
    SETTINGS.fixed = {};
    SETTINGS.manual = {};
    renderFixedGrid();
    saveSettings();
  });

  $("subjectGrid").addEventListener("click", (event) => {
    const td = event.target.closest("td[data-key]");
    if (!td || !currentSubject) return;
    const key = td.dataset.key;
    const list = SETTINGS.subjectConstraints[currentSubject] || (SETTINGS.subjectConstraints[currentSubject] = []);
    const index = list.indexOf(key);
    if (index >= 0) list.splice(index, 1);
    else list.push(key);
    renderSubjectGrid();
    saveSettings();
  });

  $("subjectClearSelected").addEventListener("click", () => {
    if (!currentSubject) return;
    delete SETTINGS.subjectConstraints[currentSubject];
    renderSubjectGrid();
    saveSettings();
  });

  $("subjectClearEvery").addEventListener("click", () => {
    SETTINGS.subjectConstraints = {};
    renderSubjectGrid();
    saveSettings();
  });

  $("teacherClearSelected").addEventListener("click", () => {
    if (!currentTeacher) return;
    delete SETTINGS.teacherUnavailable[currentTeacher];
    renderTeacherGrid();
    saveSettings();
  });

  $("teacherClearEvery").addEventListener("click", () => {
    SETTINGS.teacherUnavailable = {};
    renderTeacherGrid();
    saveSettings();
  });

  $("teacherGrid").addEventListener("click", (event) => {
    const td = event.target.closest("td[data-key]");
    if (!td || !currentTeacher) return;
    const key = td.dataset.key;
    const list = SETTINGS.teacherUnavailable[currentTeacher] || (SETTINGS.teacherUnavailable[currentTeacher] = []);
    const index = list.indexOf(key);
    if (index >= 0) list.splice(index, 1);
    else list.push(key);
    renderTeacherGrid();
    saveSettings();
  });

  $("blockWarnContinue").addEventListener("click", () => {
    $("blockWarnModal").classList.remove("show");
    pendingTab = null;
  });
  $("blockWarnSkip").addEventListener("click", () => {
    $("blockWarnModal").classList.remove("show");
    if (pendingTab) switchTab(pendingTab, true);
    pendingTab = null;
  });

  const closeDragConfirm = () => {
    $("dragConfirmModal").classList.remove("show");
    pendingDrag = null;
  };
  $("dragConfirmYes").addEventListener("click", () => {
    if (pendingDrag) {
      applyDragMove(pendingDrag.sourceTd, pendingDrag.targetTd);
      renderResultClassGrid();
      renderReport();
      saveSettings();
    }
    closeDragConfirm();
  });
  $("dragConfirmNo").addEventListener("click", closeDragConfirm);
  $("dragConfirmCancel").addEventListener("click", closeDragConfirm);
  $("dragConfirmModal").addEventListener("click", (event) => {
    if (event.target === $("dragConfirmModal")) closeDragConfirm();
  });

  $("fixedModalClose").addEventListener("click", closeFixedModal);
  $("fixedModal").addEventListener("click", (event) => {
    if (event.target === $("fixedModal")) closeFixedModal();
  });
  $("fixedModalClear").addEventListener("click", () => {
    const cls = currentFixedClass;
    const key = $("fixedModal").dataset.key;
    if (cls && key) {
      if (SETTINGS.fixed[cls]) delete SETTINGS.fixed[cls][key];
      deleteManual(cls, key);
      saveSettings();
    }
    closeFixedModal();
    renderFixedGrid();
  });
  $("fixedModalList").addEventListener("click", (event) => {
    if (event.target.closest('input[type="checkbox"]')) return;
    const option = event.target.closest(".course-option");
    if (!option) return;
    const subject = option.dataset.subject;
    const gradeChecked = option.querySelector('input[data-scope="grade"]').checked;
    const allChecked = option.querySelector('input[data-scope="all"]').checked;
    const scope = allChecked ? "all" : gradeChecked ? "grade" : "class";
    const applied = applyFixedModalChoice(subject, scope);
    $("fixedCount").textContent = applied > 0 ? `已设置 ${applied} 个班` : $("fixedCount").textContent;
  });

  $("dataUpload").addEventListener("change", async (event) => {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      showToast("请上传 .xlsx 格式的模板文件");
      event.target.value = "";
      return;
    }
    if (file.size > 1024 * 1024) {
      showToast("文件大小不能超过 1MB");
      event.target.value = "";
      return;
    }
    const buffer = await file.arrayBuffer();
    const binary = new Uint8Array(buffer).reduce((acc, byte) => acc + String.fromCharCode(byte), "");
    const contentBase64 = btoa(binary);
    const payload = await apiPost("/api/parse", { filename: file.name, contentBase64 });
    if (!payload.ok) {
      showErrors([payload.error || "上传失败"]);
      return;
    }
    writeBackup({
      filename: file.name,
      contentBase64,
      data: payload.data,
      settings: payload.settings || {},
    });
    location.reload();
  });

  $("downloadTemplate").addEventListener("click", () => {
    window.location.href = "/api/template";
  });

  $("resetData").addEventListener("click", async () => {
    clearBackup();
    await apiPost("/api/reset-data", {});
    location.reload();
  });

  $("classTitleInput").addEventListener("input", () => {
    SETTINGS.titles.classSuffix = $("classTitleInput").value.trim();
    saveSettings();
  });
  $("teacherTitleInput").addEventListener("input", () => {
    SETTINGS.titles.teacherSuffix = $("teacherTitleInput").value.trim();
    saveSettings();
  });

  $("saveRules").addEventListener("click", () => {
    SETTINGS.rules.dailyMax = Math.max(1, Math.min(12, Number($("dailyMaxInput").value) || 5));
    SETTINGS.rules.consecutiveMax = Math.max(1, Math.min(12, Number($("consecutiveMaxInput").value) || 2));
    applyPreferredSelection();
    $("ruleStatus").textContent = "规则已保存";
    saveSettings();
  });

  $("applyPreferredConsecutive").addEventListener("click", () => {
    applyPreferredSelection();
    $("ruleStatus").textContent = "优先连堂规则已保存";
    saveSettings();
  });

  $("solveButton").addEventListener("click", runSolve);
  $("resetResultButton").addEventListener("click", resetResult);
  $("exportButton").addEventListener("click", exportWorkbook);

  const classGrid = $("resultClassGrid");
  classGrid.addEventListener("dragstart", (event) => {
    const td = event.target.closest("td[data-key]");
    if (!td || td.dataset.blocked === "1" || !td.dataset.subject || td.dataset.source === "fixed") {
      event.preventDefault();
      return;
    }
    dragSource = td;
    td.classList.add("dragging");
    event.dataTransfer.effectAllowed = "move";
  });
  classGrid.addEventListener("dragover", (event) => {
    const td = event.target.closest("td[data-key]");
    if (!td || td.dataset.blocked === "1" || td.dataset.source === "fixed") return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "move";
    td.classList.add("drag-over");
  });
  classGrid.addEventListener("dragleave", (event) => {
    const td = event.target.closest("td[data-key]");
    if (td) td.classList.remove("drag-over");
  });
  classGrid.addEventListener("drop", (event) => {
    event.preventDefault();
    const td = event.target.closest("td[data-key]");
    if (td) td.classList.remove("drag-over");
    if (!dragSource) return;
    startDragMove(dragSource, td);
    dragSource = null;
  });
  classGrid.addEventListener("dragend", () => {
    classGrid.querySelectorAll(".dragging, .drag-over").forEach((cell) => cell.classList.remove("dragging", "drag-over"));
    dragSource = null;
  });
  classGrid.addEventListener("dblclick", (event) => {
    const td = event.target.closest("td[data-key]");
    if (!td || !currentResultClass || td.dataset.source !== "manual") return;
    deleteManual(currentResultClass, td.dataset.key);
    delete RESULT.schedule[currentResultClass][td.dataset.key];
    renderResultClassGrid();
    renderReport();
    saveSettings();
  });
}

function switchTab(tab, bypass = false) {
  const stepOrder = ["overview", "block", "fixed", "subject", "teacher", "rules", "result"];
  const current = document.querySelector("#tabs .tab.active")?.dataset.tab;
  if (
    !bypass &&
    current === "block" &&
    stepOrder.indexOf(tab) > stepOrder.indexOf("block")
  ) {
    const insufficient = insufficientBlockClasses();
    if (insufficient.length) {
      pendingTab = tab;
      $("blockWarnText").innerHTML =
        `<div>有以下班级不排课量不足，会出现空格：${esc(insufficient.join("、"))}</div>` +
        `<div>建议通过应用到全年级或全部班级批量设置。</div>` +
        `<div>确定下一步吗？</div>`;
      $("blockWarnModal").classList.add("show");
      return;
    }
  }
  document.querySelectorAll("#tabs .tab").forEach((button) => button.classList.toggle("active", button.dataset.tab === tab));
  document.querySelectorAll("main .panel").forEach((panel) => panel.classList.toggle("active", panel.id === tab));
}

function switchResultTab(view) {
  document.querySelectorAll("#resultTabs .result-tab").forEach((button) =>
    button.classList.toggle("active", button.dataset.view === view)
  );
  document.querySelectorAll("#resultView > div").forEach((div) => div.classList.add("hidden"));
  $(`view-${view}`).classList.remove("hidden");
}

document.addEventListener("DOMContentLoaded", init);
