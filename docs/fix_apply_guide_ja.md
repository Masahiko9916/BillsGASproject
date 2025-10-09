# 修正適用手順（初心者向け）

このドキュメントでは、`Batch sheet updates with extended last-updated helper`で提案されたスプレッドシート更新バッチ化の修正を、GitHub と Google Apps Script (GAS) に慣れていない方でも適用できるよう、順を追って説明します。

## 1. 事前準備

1. **Google アカウント**と対象の Apps Script プロジェクトへの編集権限を用意します。
2. パソコンに以下のソフトをインストールします。
   - [Git](https://git-scm.com/)（バージョン管理ツール）
   - [Node.js](https://nodejs.org/)（clasp を使う場合に必要）
3. ターミナル（Windows は PowerShell / コマンドプロンプト、macOS はターミナル）を開けるようにします。

> **メモ**: Google Apps Script エディタ上で直接ファイルを編集するだけでも適用できますが、GitHub 上の履歴と同期するために Git + clasp を使う方法を推奨します。

## 2. リポジトリを取得する

1. GitHub でリポジトリページを開き、右上の **Fork** ボタンを押して自分のアカウントにコピーを作成します（直接 push できる権限がある場合はこのステップは不要です）。
2. ターミナルで作業用フォルダに移動し、以下のコマンドを実行してローカルにリポジトリを複製します。
   ```bash
   git clone https://github.com/<あなたのアカウント名>/BillsGASproject.git
   cd BillsGASproject
   ```
3. 既にリポジトリを clone 済みであれば、最新状態を取り込むために `git pull` を実行します。

### 補足: リポジトリを常に最新に保つコツ

リモート側で他のメンバーが変更を加えている場合は、作業を始める前にローカル環境を最新の状態へ同期しましょう。初心者の方は以下の手順を順番に進めるだけで大丈夫です。

1. メインブランチ（多くの場合は `main`）に切り替えます。
   ```bash
   git switch main
   ```
2. GitHub 上の最新履歴を取得します。
   ```bash
   git fetch origin
   ```
3. 取得した最新の変更を自分の `main` に反映します。
   ```bash
   git pull --ff-only origin main
   ```
   - `--ff-only` を付けると、履歴がすでに追従している場合のみ早送り（fast-forward）で更新され、意図しないマージコミットが増えません。
4. 自分が作業するブランチに戻り、最新の `main` を取り込みます。
   ```bash
   git switch feature/apply-batch-update   # 作業ブランチ名の例
   git merge main
   ```
   ここでコンフリクト（競合）が出た場合は、Git が教えてくれる印を参考に編集してから `git add` → `git commit` で解決します。

> **豆知識**: 公式リポジトリ（例: `upstream`）から更新を取り込みたい場合は、`git remote add upstream <URL>` で upstream を登録しておき、`git fetch upstream` → `git merge upstream/main` の順で反映できます。

## 3. 作業用ブランチを作る

1. メインブランチが最新か確認します。
   ```bash
   git checkout main
   git pull
   ```
2. 修正適用用のブランチを作成して切り替えます。
   ```bash
   git switch -c feature/apply-batch-update
   ```

## 4. 修正内容を取り込む

### パターンA: 既存コミットを取り込む

既にリポジトリに修正コミットが存在する場合は、次のいずれかの方法で取り込みます。

- **マージ / チェリーピック**: 変更が入っているブランチ（例: `work`）を取り込みます。
  ```bash
  git fetch origin
  git merge origin/work
  ```
  もしくは対象コミットのハッシュが分かる場合は
  ```bash
  git cherry-pick <コミットID>
  ```

### パターンB: 手動でファイルを更新する

GitHub のプルリクエスト画面で `Files changed` を参照し、以下のファイルを Apps Script エディタまたはローカルのエディタで編集します。

- `SheetUtils.gs V4.js`
- `Main.gs V4.1.js`
- `ScheduledTasks.gs V4.1.js`
- `RegularCollection.gs V4.2.js`

アプリケーションの主要な変更点は以下です。

- `setLastUpdated_` が複数列の値をまとめて受け取れるようになりました。
- ステータス更新やイベント ID 書き込みが 1 回の範囲書き込みで完結するようにヘルパー関数を追加しています。

ローカルで編集した場合は、変更を確認するために `git status` と `git diff` を確認します。

```bash
git status
git diff
```

## 5. Apps Script と同期する

1. まだ clasp を初期化していない場合は、Google アカウントにログインします。
   ```bash
   npm install -g @google/clasp
   clasp login
   ```
2. プロジェクトディレクトリで Apps Script プロジェクト ID を設定します。
   ```bash
   clasp clone <SCRIPT_ID>
   ```
   すでに `appsscript.json` が存在する場合は `clasp pull` で最新を取得し、競合がないか確認します。
3. 編集したファイルを Apps Script に反映します。
   ```bash
   clasp push
   ```
4. GAS エディタ上で変更が反映されているか確認し、必要であればテスト実行します。

## 6. GitHub にコミット・プッシュする

1. 変更をステージングしてコミットします。
   ```bash
   git add .
   git commit -m "Apply batch sheet update helpers"
   ```
2. 自分の GitHub リポジトリにブランチをプッシュします。
   ```bash
   git push origin feature/apply-batch-update
   ```

## 7. プルリクエストを作成する

1. GitHub で新しいブランチのページを開き、`Compare & pull request` をクリックします。
2. 変更点の概要とテスト状況を記載してプルリクエストを送信します。
3. レビューアから指摘があれば対応し、承認後にマージします。

## 8. 運用時のポイント

- **スプレッドシートの権限**: 複数人で運用する場合、該当シートへの編集権限があるか事前に確認してください。
- **Apps Script の実行回数**: バッチ化によって呼び出し回数は減りますが、トリガーや手動実行が集中すると制限に達する可能性があるため、実運用前にテストシートで負荷を確認してください。
- **バックアップ**: 大きな変更を行う前に、スプレッドシートをコピーしてバックアップを作成することを推奨します。

---

ご不明点があれば、どのステップで躓いているかを具体的に書いていただければサポートしやすくなります。
