# VBAモジュール構成

## 全体構成

```text
ImportReceivingData
  ↓
M02_ExternalData
  ↓
M03_DateLogic
  ↓
M04_ReferenceCode
  ↓
M05_ReportUpdate
  ↓
作業編成画面 / 作業編成画面A

PrintIndividualInstructions
  ↓
M06_InstructionPrint
  ↓
指示書

PrintWorkAssignmentSheet
  ↓
M07_WorkAssignmentPrint
```

M08_ProtectionとM90_Commonは各処理から共通利用する。

## モジュール一覧

| モジュール | 主な役割 |
|---|---|
| M00_Constants | シート名、列番号、行番号、文言等の共通定数 |
| M01_Main | 利用者が実行する3つの公開マクロと全体制御 |
| M02_ExternalData | 外部ファイル読込、納品日・店舗番号の検証 |
| M03_DateLogic | 処理基準日、次回納品日、半月内納品回数の判定 |
| M04_ReferenceCode | 参照CD生成、重複排除、マスタ照合 |
| M05_ReportUpdate | 入荷データと作業編成画面の更新 |
| M06_InstructionPrint | 作業者別指示書の作成・印刷 |
| M07_WorkAssignmentPrint | 作業編成表の印刷 |
| M08_Protection | シート・ブック保護状態の管理 |
| M90_Common | Application状態、辞書、最終行取得等の共通処理 |
| ThisWorkbook | 起動時保護設定 |

## 公開マクロ

```vba
ImportReceivingData
PrintIndividualInstructions
PrintWorkAssignmentSheet
```

## M05の日付表示方針

作業編成画面と作業編成画面Aの日付表示は、シート数式に一本化している。

- A1: 店舗コード
- E2: 納品日
- G2: 曜日

`SetWorkAssignmentDate`、`ClearWorkAssignmentDate`は使用しない。

Excel 2019 / 2024の両方で結合セルE2:F2へのVBA操作を避けるため、VBAから値設定・表示形式設定・クリアを行わない。
