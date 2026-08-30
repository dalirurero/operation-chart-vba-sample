Attribute VB_Name = "M00_Constants"

Option Explicit



'============================================================

' GitHub公開用サンプル 共通定数

'============================================================



'--- シート名 ---

Public Const SH_MF As String = "MF"

Public Const SH_WORK As String = "作業編成画面"

Public Const SH_WORK_NEXT As String = "作業編成画面（明日）"

Public Const SH_INSTRUCTION As String = "指示書"

Public Const SH_AREA As String = "エリアデータ"

Public Const SH_IMPORT As String = "入荷データ"

Public Const SH_IMPORT_NEXT As String = "入荷データA"



'--- 外部ファイル（公開サンプルでは.xlsxを使用） ---

Public Const FILE_SEND As String = "SEND_SAMPLE.xlsx"

Public Const FILE_SEND3 As String = "SEND3_SAMPLE.xlsx"

Public Const EXTERNAL_SHEET As String = "Sheet1"



'--- MF ---

Public Const CELL_BASE_DATE As String = "D5"

Public Const CELL_DATE_OFFSET As String = "D6"

Public Const CELL_BASE_WEEKDAY As String = "E5"



'--- 行番号 ---

Public Const ROW_EXTERNAL_HEADER As Long = 1

Public Const ROW_EXTERNAL_DATA As Long = 2

Public Const ROW_IMPORT_DATA As Long = 2

Public Const ROW_WORK_HEADER As Long = 4

Public Const ROW_PRODUCT_TEMPLATE As Long = 5

Public Const ROW_WORKER_START As Long = 7

Public Const ROW_WORKER_END As Long = 36

Public Const ROW_INSTRUCTION_START As Long = 3

Public Const ROW_INSTRUCTION_FIRST_END As Long = 35

Public Const ROW_INSTRUCTION_EXTEND_START As Long = 36

Public Const ROW_WORK_PRINT_MIN As Long = 36

Public Const ROW_WORK_DEFAULT_END As Long = 51

Public Const ROW_WORK_NEXT_DEFAULT_END As Long = 73



'--- 列番号 ---

Public Const COL_REF As Long = 1             'A

Public Const COL_AREA_NAME As Long = 2       'B

Public Const COL_TOTAL_QTY As Long = 3       'C

Public Const COL_REPLACE_QTY As Long = 4     'D

Public Const COL_NO_REPLACE_QTY As Long = 5  'E

Public Const COL_REPLACE_TIME As Long = 6    'F

Public Const COL_NO_REPLACE_TIME As Long = 7 'G

Public Const COL_PLAN_TIME As Long = 8       'H

Public Const COL_ASSIGNEE As Long = 9        'I

Public Const COL_WORKER As Long = 11         'K

Public Const COL_WORKER_QTY As Long = 12     'L

Public Const COL_WORKER_TIME As Long = 13    'M



'--- 入荷データ列 ---

Public Const COL_IMPORT_REF As Long = 1       'A

Public Const COL_IMPORT_STORE As Long = 2     'B

Public Const COL_IMPORT_DATE As Long = 3      'C

Public Const COL_IMPORT_QTY As Long = 12      'L

Public Const COL_IMPORT_FLAG As Long = 15     'O



'--- エリアデータ列 ---

Public Const COL_AREA_REF As Long = 1         'A

Public Const COL_AREA_NAME_MASTER As Long = 7 'G

Public Const COL_AREA_COEFF_1 As Long = 8     'H

Public Const COL_AREA_COEFF_2 As Long = 9     'I

Public Const COL_AREA_COEFF_3 As Long = 10    'J

Public Const COL_AREA_COEFF_4 As Long = 11    'K

Public Const COL_AREA_COEFF_5 As Long = 12    'L



'--- 外部データ元列 ---

Public Const SRC_COL_STORE As Long = 1        'A

Public Const SRC_COL_DATE As Long = 2         'B

Public Const SRC_COL_SHELF1 As Long = 5       'E

Public Const SRC_COL_SHELF2 As Long = 6       'F

Public Const SRC_COL_SHELF3 As Long = 7       'G

Public Const SRC_COL_AREA As Long = 8         'H

Public Const SRC_COL_QTY As Long = 11         'K

Public Const SRC_COL_REPLACE_FLAG As Long = 14 'N（SENDのみ）

Public Const SRC_COL_SEND_LAST As Long = 14

Public Const SRC_COL_SEND3_LAST As Long = 13



'--- 表示・文言 ---

Public Const DATE_FORMAT_DISPLAY As String = "yyyy/m/d"

Public Const TEXT_SAME_DAY_REPLENISH As String = "当日補充"

Public Const VALIDATION_WORKERS As String = "=$K$5:$K$36"

Public Const PRINT_TITLE_WORK As String = "$1:$4"

Public Const PRINT_TITLE_INSTRUCTION As String = "$1:$2"



'--- 判定ステータス ---

Public Const STATUS_VALID As String = "有効"

Public Const STATUS_NO_FILE As String = "ファイルなし"

Public Const STATUS_NO_DATA As String = "データなし"

Public Const STATUS_OLD As String = "古いデータ"

Public Const STATUS_WARNING As String = "警告後続行"

Public Const STATUS_SIDE_ERROR As String = "対象側エラー"

Public Const STATUS_FATAL As String = "全体中止エラー"

Public Const STATUS_ERROR As String = "エラー"

