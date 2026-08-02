# Project Portfolio Sheets

プロジェクト横断管理用のネイティブGoogleスプレッドシートを生成・更新する、公開用Apps Scriptテンプレートです。

実プロジェクトの名称、ローカルパス、Repository URLは含みません。公開リポジトリには匿名サンプルだけを置き、実データは別の非公開リポジトリからビルド時に注入する構成です。

## Google Driveで試す

- Portfolio版: [公開テンプレートを開く](https://docs.google.com/spreadsheets/d/1bQCev4HdxU6qpLsIQAqIPkZf7SXVCLwJknwIDBn-TSU/edit?usp=sharing) / [自分のGoogle Driveへコピーする](https://docs.google.com/spreadsheets/d/1bQCev4HdxU6qpLsIQAqIPkZf7SXVCLwJknwIDBn-TSU/copy)
- Single版: [公開テンプレートを開く](https://docs.google.com/spreadsheets/d/1CFKVIIpwmEBr49hhlaEwsajoBLPVZlGvLQTtV5fjT3I/edit?usp=sharing) / [自分のGoogle Driveへコピーする](https://docs.google.com/spreadsheets/d/1CFKVIIpwmEBr49hhlaEwsajoBLPVZlGvLQTtV5fjT3I/copy)

公開テンプレートは閲覧専用です。Portfolio版は匿名サンプル3件、Single版は匿名サンプル1件を収録しています。編集する場合はコピーを作成してください。

## 管理できる内容

- Dashboard: 全件、進行中、要確認、期限超過、ブロック、危険、更新漏れ、7日／30日以内のマイルストーン、カテゴリ別・状態別一覧
- Projects: プロジェクト台帳と自動健康度（健全／注意／危険）
- Tasks: プロジェクトに紐づく実行項目
- Risks & Milestones: リスクとマイルストーンの入力元
- Dependencies: プロジェクト間の依存関係とブロッカー
- Milestone Timeline: マイルストーンの予定表
- Decision Log: 意思決定と見直し履歴
- Master: 入力候補と更新期限設定

入力規則、フィルター、固定ヘッダー、条件付き書式をApps Scriptで設定します。詳細は[データモデル](docs/data-model.md)を参照してください。

## 匿名サンプルからビルド

Node.js 18以上で実行します。

```bash
npm run build
npm run check
```

`npm run build`は`examples/projects.example.json`を使い、`dist/`へApps Script一式を生成します。

## 単一プロジェクト版（PMBOK-informed MVP）

`npm run build:single`は、ポートフォリオ版と混在しない単一案件用のApps Scriptを`dist/`へ生成します。匿名seedは`examples/single-project.example.json`、入力仕様は`schema/single-project-seed.schema.json`です。

```bash
npm run build:single
npm run build:single -- --seed /path/to/private/single-project.json --out /path/to/private/dist
```

Apps Scriptでは`createSingleProjectWorkbook`を実行します。生成するシートは「ダッシュボード」「ガイド」「プロジェクト憲章」「概要スケジュール」「WBS」「ガントチャート」「RAIDログ」「ステークホルダー」「意思決定・変更ログ」「マスター」です。マイルストーンは「WBS」の種別で管理し、予算基準線・予測・実績、全体状態、最終レビューは「プロジェクト憲章」で管理します。変更要求と意思決定は「意思決定・変更ログ」に記録します。

既存のSingle版を匿名seedから再構築する場合は、対象スプレッドシートに紐づくApps Scriptから`rebuildActiveSingleProjectWorkbook`を実行します。IDを指定できる環境では`rebuildSingleProjectWorkbook(spreadsheetId)`も利用できます。旧「作業項目」シートは「WBS」へ移行します。

「WBS」は階層、基準開始／終了日、実績開始／終了日、先行タスク、進捗率、遅延理由を保持する唯一の入力元です。「概要スケジュール」はWBSの階層1とマイルストーンを最大18件まで抽出し、週次／月次、4〜18期間を切り替える関係者向けの一枚絵です。「ガントチャート」はWBSを数式参照し、日次／週次、4〜52期間を切り替えて、計画・実績・遅延・ブロック・マイルストーンを表示します。

「ガイド」にはPMBOK第8版の7つのパフォーマンス・ドメイン（ガバナンス、スコープ、スケジュール、財務、ステークホルダー、リソース、リスク）と、テーラリング・価値・品質・説明責任の運用上の対応を記載します。予算／資源の明細、独立した変更台帳、定型ステータスレポートは後続拡張として明示し、MVPでは別シートにしません。

## 非公開データからビルド

Project Seed Schemaに準拠したJSONを`--seed`で渡します。入力ファイルは公開リポジトリへコピーされません。

```bash
npm run build -- --seed /path/to/private/projects.json --out /path/to/private/dist
```

`.clasp.json.example`を参考に、生成先を`rootDir`として`clasp push`します。Apps Scriptエディタで次の関数を実行できます。

- `createProjectPortfolio`: 新しいスプレッドシートを作成
- `upgradeProjectPortfolio(spreadsheetId)`: 既存スプレッドシートへPM機能を追加
- `verifyProjectPortfolio`: 主要シート、数式、入力規則、数式エラーを検証

Apps Scriptエディタから既存シートを更新する場合は、非公開側にIDを置いた引数なしのラッパー関数を用意してください。IDは公開リポジトリへコミットしません。

```javascript
function upgradeMyPortfolio() {
  return upgradeProjectPortfolio('YOUR_PRIVATE_SPREADSHEET_ID');
}
```

## データ境界

- 公開: Apps Script、匿名サンプル、JSON Schema、ドキュメント、検証コード
- 非公開: 実プロジェクト名、絶対パス、Repository URL、実シートID、生成済みseed

`npm run check`は公開ファイルを走査し、実データ由来の絶対パスや既知の識別子が混入していないことも確認します。
