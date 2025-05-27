# GoogleAppsScript
GASコード

## 3行でわかる動作
1. Gmailから「N高」を含む直近指定日数以内のメールを取得します。
2. OpenAI APIでイベント名・締切・要約・カテゴリを抽出します。
3. 重複を除いてスプレッドシートに追記します。

## 設定

`EventDeadlineList.js` では OpenAI API キーと出力先スプレッドシート ID をスクリプトプロパティから読み込みます。
Google Apps Script のスクリプトプロパティに次の 3 つを設定してください。

```
OPENAI_API_KEY    = <OpenAI の API キー>
SPREADSHEET_ID    = <スプレッドシート ID>
LOOKBACK_DAYS     = <取得対象とする日数(省略時は1)>
```

## 使い方（初心者向け）
1. Google Apps Script エディタに `EventDeadlineList.js` を追加します。
2. スクリプトプロパティに `OPENAI_API_KEY` と `SPREADSHEET_ID`、必要に応じて `LOOKBACK_DAYS` を登録します。
3. `summarizeNHighEmails` 関数を実行すると結果がスプレッドシートに書き込まれます。
4. トリガー画面で `cleanupExpiredEvents` を選び、時間主導型の毎日実行に設定すると、期限切れイベントが自動削除されます。

## カテゴリ分割の利用

`summarizeNHighEmails` では取得したイベントを通常の `イベント一覧` シートに追記すると同時に、
カテゴリ名だけのシートにも書き込みます。
この動作は `writeRowsUnique` の第 4 引数に `true` を渡すことで有効になっています。
利用できるカテゴリは **課外授業・重要/テスト・その他** の3種類です。

### API 呼び出し例

`DataAPI.js` の `doGet` は `category` パラメータを受け取り、
対応するカテゴリシートの内容を返します。

```
https://script.google.com/macros/s/AKfy.../exec?category=課外授業
```

パラメータを指定しなければ従来通り `イベント一覧` シートが返されます。

`mode=search` を指定するとシート内検索が行えます。利用できるパラメータは
`keyword` (件名または要約に含まれる文字列), `start` (開始日), `end`
(終了日), `category`, `limit` (最大取得件数) です。日付は `YYYY-MM-DD` 形式で指定します。`limit` を省略した場合は10件が返され、指定できる上限は20件です。

```
https://script.google.com/macros/s/AKfy.../exec?mode=search&keyword=締切&start=2024-06-01&end=2024-06-30&limit=10
```

この例では 2024/6/1 から 2024/6/30 までの期間に該当し、件名または要約に
「締切」を含む行だけが返されます。
`limit` を指定すると返される行数を制限できます。

## GPTs連携

このスクリプトは OpenAI GPTs からも利用できます。以下は簡易的な API スキーマです。

```yaml
paths:
  /exec:
    get:
      parameters:
        - name: mode
          description: search を指定すると検索モード
        - name: category
          description: カテゴリ (課外授業|重要/テスト|その他)
        - name: keyword
        - name: start
        - name: end
      - name: limit
        schema:
          type: integer
          default: 10
          maximum: 20
      responses:
        200:
          description: JSON { sheet, rows }
    post:
      requestBody:
        json: { sheet?, rows }
      responses:
        200:
          description: 追加結果
```

### プロンプト例

```
N高のイベントを調べたいので、重要/テストに該当する予定を検索してください。
API: https://script.google.com/macros/s/AKfy.../exec?mode=search&category=重要/テスト
```
