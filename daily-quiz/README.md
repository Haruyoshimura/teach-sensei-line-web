# 今日の1問

毎日1問だけ出題し、連続記録・正答率・単元別の成績を表示するWebページ。
苦手単元をもとに先生を提案する導線を持つ。

検証用のプロトタイプです。

## 構成

```
daily-quiz/
├── index.html      ページ本体（これ1枚で完結）
├── gas/Code.gs     バックエンド（Google Apps Script）
└── README.md
```

## 動かし方

`index.html` は **GAS_URL が空のままだとサンプルデータで動きます**。
そのまま開けば画面の確認ができます。

```js
const GAS_URL = "";   // ← ここにGASのURLを入れるとライブ動作
```

## URLパラメータ

| パラメータ | 例 | 説明 |
|---|---|---|
| `uid` | `U123` | ユーザー識別子。省略時は端末ごとの匿名IDを自動発行 |
| `grade` | `中2` | 学年 |
| `review` | `1` | 設計意図の注記を表示（社内レビュー用。既定は非表示） |

例: `?uid=U123&grade=中2`

## バックエンド（Google Apps Script）

新規スプレッドシート →「拡張機能 > Apps Script」→ `gas/Code.gs` を貼り付け →
`setup` 関数を1回実行すると、必要な4シート（questions / answers / stats / config）が作られます。

デプロイは「デプロイ > 新しいデプロイ > ウェブアプリ」。
発行された `/exec` URL を `index.html` の `GAS_URL` に貼ります。

偏差値・順位は `recalcStats` が計算します。
Apps Scriptの「トリガー」で **時間主導型 / 日付ベース / 午前3〜4時** に登録してください。
これが走らないとスコアは表示されません。

### questions シートの列

```
id / date / grade / subject / unit / question /
choice_a / choice_b / choice_c / choice_d / answer_index /
explain / teacher_initial / teacher_name / teacher_msg
```

- `answer_index` は **0始まり**（choice_a が正解なら 0）
- `date` にその問題を出す日を入れる。当日分が無ければ先頭行にフォールバック
- `teacher_msg` は必ず埋める。先生コメントが空の問題は出す意味が薄い

## 実装メモ

- **POSTのContent-Typeは `text/plain`。** `application/json` にするとCORSのプリフライト(OPTIONS)が飛び、
  GASは応答できないため必ず失敗します。
- 正解番号はページ読み込み時には送っていません。解答をPOSTした後に返します。
- スプレッドシートへの書き込みは `LockService` で直列化しています。
- 偏差値は同学年の母数が50人未満のとき非表示（`MIN_N`）。
  母数が小さいと数値が大きく振れるためです。
- GASのコードを更新したら **「新しいバージョン」でデプロイし直す**こと。
  保存だけでは公開URLに反映されません。
