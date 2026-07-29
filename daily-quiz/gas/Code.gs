/**
 * 今日の1問 — バックエンド（Google Apps Script Web App）
 *
 * スプレッドシートに以下4シートを作ってから使う（README参照）:
 *   questions / answers / stats / config
 *
 * デプロイ: 「デプロイ > 新しいデプロイ > ウェブアプリ」
 *   次のユーザーとして実行: 自分
 *   アクセスできるユーザー: 全員
 *   → 発行された /exec URL を index.html の GAS_URL に貼る
 */

const SS_ID = ''; // 空ならこのスクリプトが紐づくスプレッドシートを使う（コンテナバインドの場合）
const MIN_N = 50; // 同学年の母数がこれ未満なら偏差値を出さない
const SCORE_WINDOW = 30; // 偏差値の元になる直近の解答数

function ss_() {
  return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('シートが見つかりません: ' + name);
  return s;
}
function rows_(name) {
  const v = sheet_(name).getDataRange().getValues();
  if (v.length < 2) return [];
  const head = v[0].map(String);
  return v.slice(1).filter(r => String(r[0]).length).map(r => {
    const o = {};
    head.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ymd_(d) {
  return Utilities.formatDate(new Date(d), 'Asia/Tokyo', 'yyyy-MM-dd');
}

/* ============================================================
 *  GET: 今日の問題 ＋ そのユーザーの状態
 *  例) /exec?action=today&uid=U123&grade=中2
 * ========================================================== */
function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'today' || !p.action) {
      return json_({ ok: true, question: todayQuestion_(p.grade), me: userState_(p.uid, p.grade) });
    }
    if (p.action === 'me') {
      return json_({ ok: true, me: userState_(p.uid, p.grade) });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ============================================================
 *  POST: 解答を記録
 *  ※ index.html からは Content-Type: text/plain で送る。
 *     application/json にすると CORS のプリフライトが飛び、
 *     GAS は OPTIONS を返せないので必ず失敗する。ここが一番ハマる。
 * ========================================================== */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action !== 'answer') return json_({ ok: false, error: 'unknown action' });

    lock.waitLock(20000);

    const qs = rows_('questions');
    const q = qs.filter(x => String(x.id) === String(body.question_id))[0];
    if (!q) return json_({ ok: false, error: 'question not found' });

    const correct = Number(body.choice) === Number(q.answer_index);

    sheet_('answers').appendRow([
      new Date(),
      ymd_(new Date()),
      String(body.uid || 'anon'),
      String(body.grade || ''),
      String(body.question_id),
      String(q.subject),
      String(q.unit),
      Number(body.choice),
      correct ? 1 : 0,
      Number(body.ms || 0)
    ]);

    return json_({
      ok: true,
      correct: correct,
      answer_index: Number(q.answer_index),
      dist: choiceDist_(body.question_id),
      me: userState_(body.uid, body.grade)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ============================================================
 *  出題
 * ========================================================== */
function todayQuestion_(grade) {
  const today = ymd_(new Date());
  const qs = rows_('questions');
  let q = qs.filter(x => ymd_(x.date) === today && (!x.grade || !grade || x.grade === grade))[0];
  if (!q) q = qs.filter(x => !x.grade || !grade || x.grade === grade)[0]; // 未設定日はフォールバック
  if (!q) return null;
  return {
    id: String(q.id),
    subject: q.subject, unit: q.unit, q: q.question,
    choices: [q.choice_a, q.choice_b, q.choice_c, q.choice_d].filter(c => String(c).length),
    explain: q.explain,
    teacher: { i: q.teacher_initial, n: q.teacher_name, m: q.teacher_msg }
    // answer_index はここでは返さない（DevToolsで見えてしまうため）
  };
}

/** その問題の選択率（%） */
function choiceDist_(qid) {
  const a = rows_('answers').filter(r => String(r.question_id) === String(qid));
  const c = [0, 0, 0, 0];
  a.forEach(r => { const i = Number(r.choice); if (i >= 0 && i < 4) c[i]++; });
  const t = a.length || 1;
  return c.map(x => Math.round(x / t * 100));
}

/* ============================================================
 *  ユーザーの状態（連続日数・正答率・単元別・偏差値）
 * ========================================================== */
function userState_(uid, grade) {
  if (!uid) return null;
  const all = rows_('answers').filter(r => String(r.uid) === String(uid));
  if (!all.length) return { streak: 0, total: 0, accuracy: null, units: [], score: null };

  // --- 連続日数 ---
  const days = {};
  all.forEach(r => days[String(r.ymd)] = true);
  let streak = 0;
  const d = new Date();
  if (!days[ymd_(d)]) d.setDate(d.getDate() - 1); // 今日まだなら昨日から数える
  while (days[ymd_(d)]) { streak++; d.setDate(d.getDate() - 1); }

  // --- 正答率 ---
  const correct = all.filter(r => Number(r.is_correct) === 1).length;
  const accuracy = Math.round(correct / all.length * 100);

  // --- 先月との差（先週/先月の自分） ---
  const cut = new Date(); cut.setDate(cut.getDate() - 30);
  const older = all.filter(r => new Date(r.timestamp) < cut);
  const prevAcc = older.length >= 5
    ? Math.round(older.filter(r => Number(r.is_correct) === 1).length / older.length * 100)
    : null;

  // --- 単元別 ---
  const byUnit = {};
  all.forEach(r => {
    const u = String(r.unit);
    byUnit[u] = byUnit[u] || { unit: u, n: 0, ok: 0 };
    byUnit[u].n++;
    if (Number(r.is_correct) === 1) byUnit[u].ok++;
  });
  const units = Object.keys(byUnit).map(u => {
    const x = byUnit[u];
    return { unit: u, n: x.n, rate: Math.round(x.ok / x.n * 100) };
  }).sort((a, b) => a.rate - b.rate);

  // --- 直近の苦手単元（予約導線の発火判定）--- 同一単元3ミスで発火
  const recent = all.slice(-20);
  const miss = {};
  recent.forEach(r => { if (Number(r.is_correct) === 0) miss[r.unit] = (miss[r.unit] || 0) + 1; });
  let weakUnit = null;
  Object.keys(miss).forEach(u => { if (miss[u] >= 3 && !weakUnit) weakUnit = { unit: u, misses: miss[u] }; });

  return {
    streak: streak,
    total: all.length,
    accuracy: accuracy,
    accuracyDelta: prevAcc === null ? null : accuracy - prevAcc,
    units: units,
    weakUnit: weakUnit,
    score: readScore_(uid, grade)
  };
}

/** statsシートから偏差値を読む（日次バッチが書いた値） */
function readScore_(uid, grade) {
  const s = rows_('stats').filter(r => String(r.uid) === String(uid))[0];
  if (!s) return null;
  if (Number(s.n) < MIN_N) {
    return { visible: false, n: Number(s.n), minN: MIN_N }; // 母数不足は数字を出さない
  }
  return {
    visible: true,
    value: Number(s.score),
    delta: s.delta === '' ? null : Number(s.delta),
    rank: Number(s.rank),
    n: Number(s.n),
    grade: s.grade
  };
}

/* ============================================================
 *  日次バッチ：偏差値を再計算して stats シートに書く
 *  トリガー設定：時間主導型 > 日付ベース > 午前3〜4時
 * ========================================================== */
function recalcStats() {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const all = rows_('answers');

    // uid ごとに直近 SCORE_WINDOW 問の正答率を出す
    const byUid = {};
    all.forEach(r => {
      const u = String(r.uid);
      byUid[u] = byUid[u] || { uid: u, grade: String(r.grade || ''), list: [] };
      byUid[u].grade = String(r.grade || byUid[u].grade);
      byUid[u].list.push(r);
    });

    const users = Object.keys(byUid).map(u => {
      const x = byUid[u];
      const recent = x.list.slice(-SCORE_WINDOW);
      const rate = recent.filter(r => Number(r.is_correct) === 1).length / recent.length * 100;
      return { uid: x.uid, grade: x.grade, rate: rate, n: recent.length };
    }).filter(u => u.n >= 5); // 5問未満はスコア対象外

    // 学年ごとに平均・標準偏差 → 偏差値
    const grades = {};
    users.forEach(u => { (grades[u.grade] = grades[u.grade] || []).push(u); });

    const prev = {};
    rows_('stats').forEach(r => prev[String(r.uid)] = Number(r.score));

    const out = [['uid', 'grade', 'score', 'delta', 'rank', 'n', 'updated_at']];
    Object.keys(grades).forEach(g => {
      const list = grades[g];
      const mean = list.reduce((s, u) => s + u.rate, 0) / list.length;
      const sd = Math.sqrt(list.reduce((s, u) => s + Math.pow(u.rate - mean, 2), 0) / list.length) || 1;
      list.forEach(u => { u.score = Math.round((50 + 10 * (u.rate - mean) / sd) * 10) / 10; });
      list.sort((a, b) => b.score - a.score);
      list.forEach((u, i) => {
        const d = prev[u.uid] === undefined ? '' : Math.round((u.score - prev[u.uid]) * 10) / 10;
        out.push([u.uid, g, u.score, d, i + 1, list.length, new Date()]);
      });
    });

    const sh = sheet_('stats');
    sh.clear();
    sh.getRange(1, 1, out.length, out[0].length).setValues(out);
    Logger.log('recalcStats: ' + (out.length - 1) + ' users');
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ============================================================
 *  初期セットアップ（1回だけ手動実行）
 * ========================================================== */
function setup() {
  const ss = ss_();
  const defs = {
    questions: ['id', 'date', 'grade', 'subject', 'unit', 'question', 'choice_a', 'choice_b', 'choice_c', 'choice_d', 'answer_index', 'explain', 'teacher_initial', 'teacher_name', 'teacher_msg'],
    answers: ['timestamp', 'ymd', 'uid', 'grade', 'question_id', 'subject', 'unit', 'choice', 'is_correct', 'ms'],
    stats: ['uid', 'grade', 'score', 'delta', 'rank', 'n', 'updated_at'],
    config: ['key', 'value']
  };
  Object.keys(defs).forEach(name => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, defs[name].length).setValues([defs[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });

  // サンプル問題（1件だけ入れておく）
  const q = ss.getSheetByName('questions');
  if (q.getLastRow() === 1) {
    q.appendRow([
      'q001', new Date(), '中2', '数学', '二次関数',
      'y = 2x² のグラフ上で、x が 1 から 3 まで変わるときの変化の割合は？',
      '4', '6', '8', '12', 2,
      '変化の割合 =（yの増加量）÷（xの増加量）。x=1→2、x=3→18。(18−2)÷(3−1)=8。',
      '佐', '佐藤先生（数学）',
      '「2乗だから割合も2乗」って思っちゃった人、めちゃくちゃ多いところ。ここは公式より“実際に代入して引き算”が一番速いよ。'
    ]);
  }
  SpreadsheetApp.getUi && Logger.log('setup 完了');
}
