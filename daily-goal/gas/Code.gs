/**
 * 短期ゴールページ バックエンド
 * スプレッドシート: users / goals / answers / events / reports / questions
 *
 * ▼ 使い方
 *  1. 新規スプレッドシート →「拡張機能 > Apps Script」→ このファイルを貼り付け
 *  2. setup() を1回実行（6シートが作られます）
 *  3. questions シートに questions_daily_100問_*.csv をインポート
 *  4.「デプロイ > 新しいデプロイ > ウェブアプリ」→ アクセスできるユーザー: 全員
 *  5. 発行された /exec の URL を index.html の GAS_URL に貼る
 *
 * ▼ ユーザーが混ざらない設計
 *  - 端末ではなく「メールアドレス」で人を決める
 *  - メールは users シートの1箇所にだけ持つ。内部IDの uid を発行して、
 *    answers / events / goals には uid だけを書く（個人情報を散らさない）
 *  - 同じメールなら別端末でも同じ uid に戻る
 *  - 同じ端末で別のメールを入れたら別の uid になる（家族で共有しても混ざらない）
 */

const SS_ID = '';   // 空ならこのスクリプトが紐づくスプレッドシート

function ss_() { return SS_ID ? SpreadsheetApp.openById(SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { return ss_().getSheetByName(name); }
function rows_(name) {
  const sh = sh_(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const head = v.shift();
  return v.map(r => { const o = {}; head.forEach((h, i) => o[h] = r[i]); return o; });
}
function now_() { return new Date(); }
function ymd_(d) { return Utilities.formatDate(d || now_(), 'Asia/Tokyo', 'yyyy-MM-dd'); }
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ==========================================================
 *  初期セットアップ（1回だけ手動実行）
 * ========================================================== */
function setup() {
  const ss = ss_();
  const defs = {
    // メールを持つのはここだけ
    users:   ['uid', 'email', 'created_at', 'last_seen', 'note'],
    goals:   ['uid', 'goal_type', 'goal_name', 'target_date', 'unit', 'n_needed', 'n_done', 'updated_at'],
    answers: ['timestamp', 'ymd', 'uid', 'question_id', 'guess', 'choice', 'is_correct', 'ms'],
    events:  ['timestamp', 'ymd', 'uid', 'event', 'detail'],
    reports: ['timestamp', 'uid', 'question_id', 'body', 'status'],
    questions: ['id', 'tier', 'date', 'grade', 'order', 'subject', 'unit', 'source',
                'question', 'choice_a', 'choice_b', 'choice_c', 'choice_d', 'answer_index',
                'hint', 'explain', 'unit_tag', 'bridge_unit',
                'teacher_id', 'teacher_initial', 'teacher_name', 'teacher_intro', 'teacher_msg']
  };
  Object.keys(defs).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, defs[name].length).setValues([defs[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  Logger.log('setup 完了。questions シートにCSVをインポートしてください。');
}

/* ==========================================================
 *  ユーザー（メール → uid）
 * ========================================================== */
function normEmail_(e) { return String(e || '').trim().toLowerCase(); }

function findOrCreateUser_(email) {
  const em = normEmail_(email);
  if (!/^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$/.test(em)) return null;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sh_('users');
    const list = rows_('users');
    const hit = list.filter(u => normEmail_(u.email) === em)[0];
    if (hit) {
      // last_seen を更新
      const idx = list.indexOf(hit) + 2;
      sh.getRange(idx, 4).setValue(now_());
      return String(hit.uid);
    }
    const uid = 'u' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    sh.appendRow([uid, em, now_(), now_(), '']);
    return uid;
  } finally {
    lock.releaseLock();
  }
}

/* ==========================================================
 *  今日の1問
 * ========================================================== */
function todayQuestion_() {
  const t = ymd_();
  const qs = rows_('questions').filter(q => String(q.tier) === 'daily');
  // date が今日のもの。無ければ、過ぎている中でいちばん新しいもの
  let q = qs.filter(x => ymd_(new Date(x.date)) === t)[0];
  if (!q) {
    const past = qs.filter(x => ymd_(new Date(x.date)) <= t)
                   .sort((a, b) => new Date(b.date) - new Date(a.date));
    q = past[0];
  }
  return q || null;
}

/**
 * 正答率の出し方（嘘の初期値は置かない）
 *   gov      … 出典のある全国調査（source に「文化庁」を含む場合）
 *   pct      … 回答30人以上。%で出す
 *   few      … 回答30人未満。人数で出す
 *   tomorrow … まだ誰も解いていない
 */
const RATE_MIN_PCT = 30;

function rateInfo_(question_id, source) {
  const a = rows_('answers').filter(r => String(r.question_id) === String(question_id));
  const n = a.length;
  const ok = a.filter(r => String(r.is_correct) === 'true' || r.is_correct === true).length;
  if (/文化庁|全国学力/.test(String(source || ''))) {
    return { mode: 'gov', n: n, ok: ok };   // 実数は questions.source 側の記述を index 側で使う
  }
  if (n === 0) return { mode: 'tomorrow', n: 0, ok: 0 };
  if (n < RATE_MIN_PCT) return { mode: 'few', n: n, ok: ok };
  return { mode: 'pct', n: n, ok: ok, rate: Math.round(ok / n * 100) };
}

/* ==========================================================
 *  Web API
 * ========================================================== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const uid = String(p.uid || '');
  try {
    if (p.action === 'today') {
      const q = todayQuestion_();
      if (!q) return json_({ ok: false, error: 'no_question' });
      const answered = uid ? rows_('answers')
        .filter(r => String(r.uid) === uid && String(r.question_id) === String(q.id))[0] : null;
      return json_({
        ok: true,
        question: {
          id: q.id, question: q.question,
          choices: [q.choice_a, q.choice_b, q.choice_c, q.choice_d],
          hint: q.hint, unit: q.unit, subject: q.subject,
          bridge_unit: q.bridge_unit, unit_tag: q.unit_tag, source: q.source
          // answer_index と explain は答えてから返す（先に見えないように）
        },
        answered: answered ? {
          choice: Number(answered.choice), guess: Number(answered.guess)
        } : null,
        rate: rateInfo_(q.id, q.source),
        goal: uid ? (rows_('goals').filter(g => String(g.uid) === uid)[0] || null) : null
      });
    }
    if (p.action === 'reveal') {   // 答え合わせ（回答済みのときだけ）
      const q = todayQuestion_();
      const done = rows_('answers')
        .filter(r => String(r.uid) === uid && String(r.question_id) === String(q.id))[0];
      if (!done) return json_({ ok: false, error: 'not_answered' });
      return json_({ ok: true, answer_index: Number(q.answer_index), explain: q.explain,
                     rate: rateInfo_(q.id, q.source) });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'bad_json' }); }
  try {
    /* ---- 初回：メールから uid を発行 ---- */
    if (b.action === 'identify') {
      const uid = findOrCreateUser_(b.email);
      if (!uid) return json_({ ok: false, error: 'bad_email' });
      logEvent_(uid, 'identify', '');
      return json_({ ok: true, uid: uid });
    }

    const uid = String(b.uid || '');
    if (!uid) return json_({ ok: false, error: 'no_uid' });

    /* ---- ゴールの設定・更新 ---- */
    if (b.action === 'save_goal') {
      const lock = LockService.getScriptLock(); lock.waitLock(10000);
      try {
        const sh = sh_('goals'), list = rows_('goals');
        const hit = list.filter(g => String(g.uid) === uid)[0];
        const row = [uid, b.goal_type || '', b.goal_name || '', b.target_date || '',
                     b.unit || '', Number(b.n_needed || 0), Number(b.n_done || 0), now_()];
        if (hit) sh.getRange(list.indexOf(hit) + 2, 1, 1, row.length).setValues([row]);
        else sh.appendRow(row);
      } finally { lock.releaseLock(); }
      logEvent_(uid, 'save_goal', b.unit || '');
      return json_({ ok: true });
    }

    /* ---- 回答（1問1回だけ） ---- */
    if (b.action === 'answer') {
      const q = todayQuestion_();
      if (!q) return json_({ ok: false, error: 'no_question' });
      const dup = rows_('answers')
        .filter(r => String(r.uid) === uid && String(r.question_id) === String(q.id))[0];
      if (!dup) {
        const isOk = Number(b.choice) === Number(q.answer_index);
        sh_('answers').appendRow([now_(), ymd_(), uid, q.id,
          Number(b.guess), Number(b.choice), isOk, Number(b.ms || 0)]);
      }
      return json_({ ok: true, answer_index: Number(q.answer_index), explain: q.explain,
                     rate: rateInfo_(q.id, q.source) });
    }

    /* ---- 画面を開いた / 予約ボタンを押した（判断指標の入口） ---- */
    if (b.action === 'event') {
      logEvent_(uid, String(b.event || ''), String(b.detail || ''));
      return json_({ ok: true });
    }

    /* ---- 「おかしいかも？」の報告 ---- */
    if (b.action === 'report') {
      sh_('reports').appendRow([now_(), uid, String(b.question_id || ''),
                                String(b.body || '').slice(0, 1000), '未対応']);
      logEvent_(uid, 'report', String(b.question_id || ''));
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function logEvent_(uid, ev, detail) {
  sh_('events').appendRow([now_(), ymd_(), uid, ev, detail]);
}

/* ==========================================================
 *  日次の集計（トリガーで毎朝まわす）
 *  スプレッドシートの「集計」シートに、その日の漏斗を1行ずつ足す
 * ========================================================== */
function dailyRollup() {
  const ss = ss_();
  let sh = ss.getSheetByName('集計');
  if (!sh) {
    sh = ss.insertSheet('集計');
    sh.getRange(1, 1, 1, 7).setValues([['ymd', 'ゴール設定者', 'ページ来訪', '回答', '予約ボタン', '報告', '来訪→予約']])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const d = ymd_(new Date(Date.now() - 86400000));   // 前日ぶん
  const ev = rows_('events').filter(r => ymd_(new Date(r.timestamp)) === d);
  const uniq = f => new Set(ev.filter(f).map(r => r.uid)).size;
  const open = uniq(r => r.event === 'open');
  const book = uniq(r => r.event === 'book_click');
  sh.appendRow([d,
    uniq(r => r.event === 'save_goal'),
    open,
    new Set(rows_('answers').filter(r => String(r.ymd) === d).map(r => r.uid)).size,
    book,
    uniq(r => r.event === 'report'),
    open ? Math.round(book / open * 100) + '%' : '']);
}
