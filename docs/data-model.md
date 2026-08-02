# データモデル

`web/project-plan-workspace`のWork Item中心の状態・期限・ブロッカー管理を、プロジェクト横断の棚卸し用途へ拡張したモデルです。

## シート構成

| シート | 役割 | 主キー・参照 |
| --- | --- | --- |
| Dashboard | 全管理シートの数式集計 | 各シートを参照 |
| Projects | プロジェクト台帳と健康度 | `Project ID` |
| Tasks | 実行項目 | `Task ID`、`Project ID` → Projects |
| Risks & Milestones | リスクとマイルストーンの入力元 | `Project ID` → Projects |
| Dependencies | プロジェクト間依存 | `Dependency ID`、2つのProject ID → Projects |
| Milestone Timeline | マイルストーンの自動予定表 | Risks & Milestonesを参照 |
| Decision Log | 意思決定履歴 | `Decision ID`、`Project ID` → Projects |
| Master | 入力候補・更新期限 | 各入力規則から参照 |

## プロジェクト健康度

Projectsの`健康度`と`健康度理由`は数式で自動判定します。

- 危険: プロジェクト、依存関係、またはリスクがブロック中
- 注意: 棚卸し待ち、期限超過、最終確認日なし、または更新期限超過
- 健全: 上記に該当しない

完了・中止したプロジェクトは更新漏れ判定から除外します。更新期限はMasterの`更新期限(日)`で変更でき、初期値は30日です。

## マイルストーン

`Risks & Milestones`で種別を`マイルストーン`にすると、`Milestone Timeline`へ期限順に自動表示されます。開始日、期限、進捗率、外部URLを管理できます。
Dashboardでは未完了マイルストーンを7日以内・30日以内に分けて数式集計します。

## データ分離

公開テンプレートは匿名サンプルのみを保持します。実データはJSON Schemaに準拠した非公開seedとして管理し、ビルド時だけApps Scriptへ合成します。
