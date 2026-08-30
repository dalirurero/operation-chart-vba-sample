# Operation Chart VBA Sample

Excel VBAで、納品データ・納品予定データの取込からマスタ照合、作業割当、個人別指示書・作業編成表の作成までを行うサンプルです。

実務で使用した業務ツールの構成をベースに、GitHub公開用としてデータ・名称を匿名化したものです。会社名、店舗名、実際の商品コード、従業員名、数量などの実データは含みません。

## 主な機能

- 納品データ・納品予定データ（外部Excelファイル）の取込
- 納品日の確認
- 参照CDの自動生成
- エリアマスタとの照合
- 参照CDの重複排除・昇順化
- 商品件数に応じた作業編成画面の更新
- 入力規則を利用した作業者割当
- 個人別指示書の作成・印刷
- 作業編成表の印刷
- シート保護を考慮した更新処理
- エラー時の中止・既存データ保護
- Excel 2019 / 2024を考慮した実装

## 処理の流れ

```text
納品データ / 納品予定データ
   ↓
データ取込
   ↓
日付確認
   ↓
参照CD生成
   ↓
エリアマスタ照合
   ↓
作業編成画面更新
   ↓
担当者設定
   ↓
個人別指示書 / 作業編成表
```

## フォルダ構成

```text
operation-chart-vba-sample/
├─ README.md
├─ LICENSE
├─ .gitignore
├─ .gitattributes
├─ src/
│  ├─ M00_Constants.bas
│  ├─ M01_Main.bas
│  ├─ M02_ExternalData.bas
│  ├─ M03_DateLogic.bas
│  ├─ M04_ReferenceCode.bas
│  ├─ M05_ReportUpdate.bas
│  ├─ M06_InstructionPrint.bas
│  ├─ M07_WorkAssignmentPrint.bas
│  ├─ M08_Protection.bas
│  ├─ M90_Common.bas
│  └─ ThisWorkbook.cls
├─ sample/
│  ├─ OperationChart_Sample.xlsm
│  ├─ SEND_SAMPLE.xlsx
│  └─ SEND3_SAMPLE.xlsx
├─ docs/
│  ├─ specification.md
│  ├─ module_overview.md
│  └─ user_manual.pdf
└─ screenshots/
   ├─ main_screen.png
   ├─ work_assignment.png
   └─ instruction_sheet.png
```


## 画面イメージ

### MF

![MF](screenshots/main_screen.png)

### 作業編成画面

![作業編成画面](screenshots/work_assignment.png)

### 指示書

![指示書](screenshots/instruction_sheet.png)

## サンプルデータ

作業者は `作業者A`、`作業者B` の2名です。商品・棚・エリア・数量はすべて架空の値です。

参照CDは、実装と同じ規則で生成します。

```text
参照CD = 棚1 × 100 + 棚2 + エリアコード
```

例：

```text
棚1 = 01
棚2 = 02
エリア = 1B
→ 1021B
```

## 動作環境

- Windows
- Microsoft Excel デスクトップ版
- Excel 2019 / Excel 2024を想定
- マクロを有効にできる環境

## セットアップ

`OperationChart_Sample.xlsm` には、動作に必要なVBAがあらかじめ組み込まれています。  
`src/` フォルダには、GitHub上でVBAコードを確認できるよう、同じVBAソースをテキストファイルとして収録しています。

1. `sample/` フォルダ内のファイルを同じフォルダに配置します。
2. Excelで `OperationChart_Sample.xlsm` を開きます。
3. マクロが無効になっている場合は、マクロを有効にします。
4. `MF` シートの「入荷データ取込」ボタンから、納品データ・納品予定データを取り込みます。
5. 作業編成画面で担当者を設定し、必要に応じて指示書・作業編成表を印刷します。

VBAを実行するために、`src/` の `.bas` / `.cls` ファイルを手動でインポートする必要はありません。

`src/` は、VBAの構成や実装内容をGitHub上で確認するために収録しています。

## 日付について

`SEND_SAMPLE.xlsx` は当日、`SEND3_SAMPLE.xlsx` はVBAの次回納品日判定と同じ規則になるよう日付を数式で設定しています。通常は日付調整値を空白のまま実行してください。

水曜日は納品がありません。当日または次回予定日に水曜日が含まれる場合、対象データが取り込まれないことがありますが、異常ではありません。

## 公開用に変更した点

- 実際の作業者名 → `作業者A`、`作業者B`
- 店舗コード・商品情報・数量 → 架空値
- エリア名 → `エリアA` などの架空名
- 外部サンプルファイル → GitHubで扱いやすい `.xlsx`
- VBAのモジュール名・プロシージャ名 → 英字化
- 実データ、会社名、社内パス等はすべて除外

シート名は実運用版の構成を示すため、日本語名を維持しています。

## 注意

本リポジトリは業務自動化の設計・実装例を示すサンプルです。実運用へ転用する場合は、データ列、マスタ構成、印刷設定、保護設定等を利用環境に合わせて確認してください。
