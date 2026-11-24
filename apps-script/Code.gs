// 스쿨서비스 파트 — 시트 저장 서버
// 스크립트 속성: SHEET_ID(시트 ID)

const TZ = 'Asia/Seoul';
const SH = { tasks: '업무리스트', records: '업무기록', requests: '확인요청', news: '파트 소식', plans: '개선과제', lists: '목록' };
const COLS = { tasks: 9, records: 9, requests: 11, news: 5, plans: 6 };
const STATUSES = ['진행중', '대기', '보류', '완료'];

function doGet() {
  return ContentService.createTextOutput('OK');
}

function doPost(e) {
  const api = { getData: getData, addTask: addTask, editTask: editTask, addRecord: addRecord, updateRequest: updateRequest, addNews: addNews };
  let out;
  try {
    const body = JSON.parse(e.postData.contents);
    if (!Object.prototype.hasOwnProperty.call(api, body.fn)) {
      out = { ok: false, error: '지원하지 않는 작업이에요.' };
    } else {
      out = { ok: true, data: api[body.fn].apply(null, Array.isArray(body.args) ? body.args : []) };
    }
  } catch (err) {
    out = { ok: false, error: (err && err.message) || '처리하지 못했어요.' };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function prop_(name) { return PropertiesService.getScriptProperties().getProperty(name) || ''; }

let SS_ = null;
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(prop_('SHEET_ID'))); }
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('시트에 ‘' + name + '’ 탭이 없어요.');
  return sh;
}
function str_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v).trim();
}
// '='로 시작하는 글이 수식이 되지 않게 한다
function text_(v) {
  const s = String(v === null || v === undefined ? '' : v).trim();
  return s.charAt(0) === '=' ? ' ' + s : s;
}
function id_(v) {
  const s = str_(v), n = Number(s);
  return s !== '' && isFinite(n) ? n : s;
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function date_(s) {
  const raw = String(s || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) throw new Error('날짜를 확인해 주세요.');
  const d = Utilities.parseDate(raw, TZ, 'yyyy-MM-dd');
  if (Utilities.formatDate(d, TZ, 'yyyy-MM-dd') !== raw) throw new Error('날짜를 확인해 주세요.');
  return d;
}

// 2행부터 내용이 있는 행을 { r: 시트 행 번호, v: 값 }으로 읽는다
function rows_(kind) {
  const sh = sheet_(SH[kind]);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, COLS[kind]).getValues()
    .map(function (v, i) { return { r: i + 2, v: v }; })
    .filter(function (x) { return x.v.some(function (c) { return str_(c) !== ''; }); });
}
function nextId_(rows) {
  return rows.reduce(function (m, x) { return Math.max(m, Number(x.v[0]) || 0); }, 0) + 1;
}
// 글자 칸은 텍스트 서식으로 두어 '11.4' 같은 글이 숫자나 날짜로 바뀌지 않게 한다
function write_(kind, row, values, textCols, dateCols) {
  const sh = sheet_(SH[kind]);
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
  textCols.forEach(function (c) { sh.getRange(row, c).setNumberFormat('@'); });
  dateCols.forEach(function (c) { sh.getRange(row, c).setNumberFormat('yyyy-mm-dd'); });
  sh.getRange(row, 1, 1, values.length).setValues([values]);
}
function append_(kind, values, textCols, dateCols) {
  const rows = rows_(kind);
  write_(kind, rows.length ? rows[rows.length - 1].r + 1 : 2, values, textCols, dateCols);
}
function column_(a1) {
  return sheet_(SH.lists).getRange(a1).getValues().map(function (r) { return str_(r[0]); }).filter(String);
}
function names_() { return column_('C2:C'); }
function cats_() { return column_('B2:B'); }

function readTasks_() {
  return rows_('tasks').filter(function (x) { return str_(x.v[2]); }).map(function (x) {
    const v = x.v;
    return { r: x.r, no: str_(v[0]), priority: Number(v[1]) || 0, name: str_(v[2]), cat: str_(v[3]), owner: str_(v[4]),
             status: str_(v[5]), due: str_(v[6]), issue: str_(v[7]), link: str_(v[8]) };
  });
}

function getData() {
  const tasks = readTasks_().map(function (t) { delete t.r; return t; });
  const records = rows_('records').map(function (x) {
    const v = x.v;
    return { row: x.r, recordId: str_(v[0]), taskId: str_(v[1]), task: str_(v[2]), date: str_(v[3]), author: str_(v[4]),
             done: str_(v[5]), next: str_(v[6]), created: str_(v[7]), from: '', to: '', req: '', reqStatus: '', reply: '' };
  }).filter(function (r) { return r.task; });
  const requests = rows_('requests').map(function (x) {
    const v = x.v;
    return { row: 1000000 + x.r, requestId: str_(v[0]), taskId: str_(v[1]), task: str_(v[2]), date: str_(v[3]), author: '',
             done: '', next: '', from: str_(v[4]), to: str_(v[5]), req: str_(v[6]), reqStatus: str_(v[7]), reply: str_(v[8]) };
  }).filter(function (r) { return r.task; });
  const news = rows_('news').filter(function (x) { return str_(x.v[1]); }).map(function (x) {
    const v = x.v;
    return { date: str_(v[0]), title: str_(v[1]), summary: str_(v[2]), link: str_(v[3]), by: str_(v[4]) };
  }).reverse().sort(function (a, b) { return b.date.localeCompare(a.date); });
  const plans = rows_('plans').filter(function (x) { return str_(x.v[1]); }).map(function (x) {
    const v = x.v;
    return { no: Number(v[0]) || 999, title: str_(v[1]), why: str_(v[2]), first: str_(v[3]), stage: str_(v[4]), memo: str_(v[5]) };
  }).sort(function (a, b) { return a.no - b.no; });
  return { today: today_(), names: names_(), cats: cats_(), tasks: tasks, records: records.concat(requests), news: news, plans: plans };
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('다른 사람이 저장하고 있어요. 잠시 뒤 다시 눌러 주세요.');
  try {
    const out = fn();
    SpreadsheetApp.flush();
    return out;
  } finally {
    lock.releaseLock();
  }
}
function member_(name) {
  const n = String(name || '').trim();
  if (names_().indexOf(n) < 0) throw new Error('이름을 확인해 주세요.');
  return n;
}

/* 업무 */
function addTask(p) { return saveTask_(p, false); }
function editTask(p) { return saveTask_(p, true); }

function rank_(t) {
  const n = Number(t.priority || t.no);
  return n > 0 ? n : Infinity;
}
function byRank_(a, b) {
  const x = rank_(a), y = rank_(b);
  return x === y ? 0 : x - y;
}

function saveTask_(p, editing) {
  p = p || {};
  const name = text_(p.name), rank = Number(p.priority);
  if (!name) throw new Error('업무명을 입력해 주세요.');
  if (!Number.isSafeInteger(rank) || rank < 1 || rank > 100000) throw new Error('우선순위를 확인해 주세요.');
  if (STATUSES.indexOf(p.status) < 0) throw new Error('상태를 선택해 주세요.');
  if (!Array.isArray(p.owners) || !p.owners.length) throw new Error('담당자를 선택해 주세요.');
  if (p.due) date_(p.due);
  const rec = p.record || {};
  const hasRecord = !!(text_(rec.done) || text_(rec.next));
  if (hasRecord) date_(rec.date);

  return withLock_(function () {
    const members = names_();
    if (!p.owners.every(function (n) { return members.indexOf(n) >= 0; })) throw new Error('담당자를 확인해 주세요.');
    const author = hasRecord ? member_(p.me) : '';
    const tasks = readTasks_();
    const existing = editing ? tasks.filter(function (t) { return t.no === String(p.id); })[0] : null;
    if (editing && !existing) throw new Error('업무를 찾지 못했어요. 새로고침해 주세요.');
    if (tasks.some(function (t) { return t.name === name && t !== existing; })) throw new Error('이미 등록된 업무명이에요.');
    if (existing && p.base && ['name', 'cat', 'owner', 'status', 'due', 'issue', 'link'].some(function (k) { return str_(p.base[k]) !== existing[k]; })) {
      throw new Error('다른 사람이 먼저 수정한 업무예요. 새로고침 후 다시 저장해 주세요.');
    }

    const completed = p.status === '완료';
    const newlyCompleted = completed && (!existing || existing.status !== '완료');
    const old = {};
    tasks.forEach(function (t) { old[t.no] = t.priority; });
    const peers = tasks.filter(function (t) { return t !== existing && (t.status === '완료') === completed; }).sort(byRank_);
    const oldName = existing ? existing.name : '';
    const task = existing || { no: String(tasks.reduce(function (m, t) { return Math.max(m, Number(t.no) || 0); }, 0) + 1), r: 0 };
    task.priority = newlyCompleted ? 1000000 : rank;
    task.name = name;
    task.cat = text_(p.cat);
    task.owner = p.owners.join(', ');
    task.status = p.status;
    task.due = p.due || '';
    task.issue = text_(p.issue);
    task.link = text_(p.link);
    if (!existing) tasks.push(task);
    if (!newlyCompleted) {
      const first = completed ? tasks.filter(function (t) { return t.status !== '완료'; }).length + 1 : 1;
      peers.splice(Math.max(0, rank - first), 0, task);
      peers.forEach(function (t, i) { t.priority = first + i; });
    }
    tasks.slice().sort(function (a, b) {
      return (a.status === '완료' ? 1 : 0) - (b.status === '완료' ? 1 : 0) || byRank_(a, b);
    }).forEach(function (t, i) { t.priority = i + 1; });

    const values = [id_(task.no), task.priority, task.name, task.cat, task.owner, task.status, task.due ? date_(task.due) : '', task.issue, task.link];
    if (existing) write_('tasks', task.r, values, [3, 8, 9], [7]);
    else append_('tasks', values, [3, 8, 9], [7]);
    const sh = sheet_(SH.tasks);
    tasks.forEach(function (t) {
      if (t !== task && t.priority !== old[t.no]) sh.getRange(t.r, 2).setValue(t.priority);
    });
    if (existing && oldName !== name) renameTask_(task.no, oldName, name);
    if (hasRecord) putRecord_(task, rec.date, author, rec.done, rec.next);
    return getData();
  });
}
function renameTask_(no, oldName, name) {
  ['records', 'requests'].forEach(function (kind) {
    const sh = sheet_(SH[kind]);
    rows_(kind).forEach(function (x) {
      const taskId = str_(x.v[1]);
      if (taskId === no || (!taskId && str_(x.v[2]) === oldName)) {
        sh.getRange(x.r, 2).setValue(id_(no));
        sh.getRange(x.r, 3).setNumberFormat('@').setValue(name);
      }
    });
  });
}

/* 기록·요청 */
function putRecord_(task, date, author, done, next) {
  append_('records', [nextId_(rows_('records')), id_(task.no), task.name, date_(date), author, text_(done), text_(next), now_(), ''],
    [3, 5, 6, 7, 8, 9], [4]);
}
function addRecord(p) {
  p = p || {};
  const done = text_(p.done), next = text_(p.next), req = text_(p.req);
  date_(p.date);
  if (!done && !next && !req) throw new Error('기록할 내용을 입력해 주세요.');
  return withLock_(function () {
    const me = member_(p.from);
    const task = readTasks_().filter(function (t) { return t.name === String(p.task || '').trim(); })[0];
    if (!task) throw new Error('업무를 찾지 못했어요. 새로고침해 주세요.');
    let to = '';
    if (req) {
      to = member_(p.to);
      if (to === me) throw new Error('요청 받는 사람을 확인해 주세요.');
    }
    if (done || next) putRecord_(task, p.date, me, done, next);
    if (req) {
      append_('requests', [nextId_(rows_('requests')), id_(task.no), task.name, date_(p.date), me, to, req, '확인 전', '', '', ''],
        [3, 5, 6, 7, 8, 9, 10, 11], [4]);
    }
    return getData();
  });
}
function updateRequest(id, status, reply) {
  if (['확인 중', '완료'].indexOf(status) < 0) throw new Error('요청 상태를 확인해 주세요.');
  return withLock_(function () {
    const found = rows_('requests').filter(function (x) { return str_(x.v[0]) === String(id); })[0];
    if (!found || str_(found.v[7]) === '완료') throw new Error('이미 처리했거나 찾을 수 없는 요청이에요. 새로고침해 주세요.');
    const answer = text_(reply);
    const replied = answer !== str_(found.v[8]) ? now_() : str_(found.v[9]);
    sheet_(SH.requests).getRange(found.r, 8, 1, 4).setNumberFormat('@')
      .setValues([[status, answer, replied, status === '완료' ? now_() : '']]);
    return getData();
  });
}

/* 소식 */
function addNews(p) {
  p = p || {};
  const title = text_(p.title), link = text_(p.link);
  if (!title) throw new Error('소식 제목을 적어 주세요.');
  if (link && !/^https?:\/\//i.test(link)) throw new Error('문서 링크를 확인해 주세요.');
  return withLock_(function () {
    append_('news', [date_(today_()), title, text_(p.summary), link, member_(p.by)], [2, 3, 4, 5], [1]);
    return getData();
  });
}
