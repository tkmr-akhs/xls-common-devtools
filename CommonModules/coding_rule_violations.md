# 共通モジュール コーディング規約違反一覧

生成日: 2026-05-13

## 対象

- 対象ディレクトリ: `CommonModules/modules`
- 対象ファイル: `.bas` / `.cls`
- 除外: `Test_*.bas` / `Test_*.cls`。ユニットテスト本体は今回の洗い出し対象外です。
- `Lib_UnitTest.bas`、`UnitTestAssert.cls`、`UnitTestUtils.cls`、`*TestDouble.cls` は `Test_` で始まらない共通モジュールとして対象に含めています。
- `CommonModules/todo.md` とは別管理です。
- この一覧は、修正せずに保留または無視しているコーディング規約違反を記録します。

## 無視するもの / 対象外

- ユニットテスト本体 (`Test_*`) の規約違反。
- 直接公開フィールドであること自体。改訂後の規約で、読み書き両方を受け入れる場合は外部公開が許容されました。
- シートから利用される大文字英数のみの公開手続き名。
- `ParamArray` 引数に `ByVal` / `ByRef` がないこと。VBA の構文制約として扱います。
- バグ、仕様改善、リファクタリング候補は `CommonModules/todo.md` 側で扱います。

## 判定メモ

- `要修正`: 規約上そのまま違反として扱えるもの。
- `要確認`: 例外や既存仕様として許容される可能性があるため、修正前に意図を確認したいもの。
- ローカル変数名と `Integer` 型の index/count 用途は機械検出を含みます。誤検出の可能性がある項目は `要確認` としています。

## サマリ

| ルール | 要修正 | 要確認 | 合計 |
| --- | ---: | ---: | ---: |
| モジュール名長 | 0 | 3 | 3 |
| 公開グローバル変数名 | 0 | 2 | 2 |
| 公開手続き名 | 1 | 0 | 1 |
| 公開引数名 | 1 | 0 | 1 |
| ファクトリ関数配置 | 0 | 2 | 2 |
| Option Explicit 欠落 | 1 | 0 | 1 |
| モジュール冒頭順序 / コメント | 6 | 0 | 6 |
| スコープ未明示のモジュール定数 | 14 | 0 | 14 |
| Private 補助手続き名 | 2 | 0 | 2 |
| ローカル変数名 | 130 | 0 | 130 |
| Integer 型の index/count 用途 | 0 | 40 | 40 |

## 残件

ここに残るものは、互換性、公開 API、テスト基盤、または広範囲の機械的リネームに影響するため保留しています。動作に影響しない明示化でも、対象範囲が広いものはこの一覧に記録します。

## 詳細

### モジュール名長 (3件)

- [無視] ApplicationScreenUpdateManager (Module) `CommonModules/modules/ApplicationScreenUpdateManager.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=30
  - 無視する理由: 利便性および互換性のため
- [無視] FileSystemServiceTestDouble (Module) `CommonModules/modules/FileSystemServiceTestDouble.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=27
  - 無視する理由: 利便性および互換性のため
- [無視] WorksheetRangeBoundsEnumerator (Module) `CommonModules/modules/WorksheetRangeBoundsEnumerator.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=30
  - 無視する理由: 利便性および互換性のため

### 公開グローバル変数名 (2件)

- [無視] Lib_FileSystem.FsSrv (Variable) `CommonModules/modules/Lib_FileSystem.bas:13`: 利用者から見える Public 変数です。基盤サービスとして公開名を許容するか確認してください。
  - 無視する理由: 正当であるため (コーディング規約の記載を見直したほうがいいかも？)
- [無視] Lib_TextFile.TfSrv (Variable) `CommonModules/modules/Lib_TextFile.bas:13`: 利用者から見える Public 変数です。基盤サービスとして公開名を許容するか確認してください。
  - 無視する理由: 正当であるため (コーディング規約の記載を見直したほうがいいかも？)

### 公開手続き名 (1件)

- [要修正] UnitTestAssert.pEqualsCore (Sub) `CommonModules/modules/UnitTestAssert.cls:149`: `Public` 手続きですが、Private 補助手続き用の `p` プレフィックス名です。

### 公開引数名 (1件)

- [要修正] Lib_Common.FormatIDName.sep (Argument) `CommonModules/modules/Lib_Common.bas:1335`: 公開手続きの引数名が PascalCase ではありません。

### ファクトリ関数配置 (2件)

- [無視] New_RangeBounds (Function) `CommonModules/modules/Lib_Common.bas:319`: 引数付きコンストラクター相当の `New_Xxx` 関数ですが、ファクトリ用の `Constructor.bas` ではなく `Lib_Common.bas` にあります。
  - 無視する理由: 正当であるため (コーディング規約の記載を見直したほうがいいかも？)
- [無視] New_InputSheet (Function) `CommonModules/modules/Lib_Common.bas:349`: 引数付きコンストラクター相当の `New_Xxx` 関数ですが、ファクトリ用の `Constructor.bas` ではなく `Lib_Common.bas` にあります。
  - 無視する理由: 正当であるため (コーディング規約の記載を見直したほうがいいかも？)

### Option Explicit 欠落 (1件)

- [要修正] Tmp_ManualTest.bas `CommonModules/modules/Tmp_ManualTest.bas:1`: `Option Explicit` がありません。
  - 無視する理由: 一時的なファイルであり後で消すため

### モジュール冒頭順序 / コメント (6件)

- [無視] ArrayObject.cls `CommonModules/modules/ArrayObject.cls:12`: `Option Explicit` の直後にモジュールコメントがなく、Private フィールド宣言が始まっています。
  - 無視する理由: Doxygen 仕様に対応する目的で、正当であるため (コーディング規約の記載を見直したほうがいいかも？)
- [無視] ObjectList.cls `CommonModules/modules/ObjectList.cls:13`: `Option Base 0` の直後にモジュールコメントではなく `Private Const` が置かれています。
  - 無視する理由: Doxygen 仕様に対応する目的で、正当であるため (コーディング規約の記載を見直したほうがいいかも？)
- [無視] Tmp_ManualTest.bas `CommonModules/modules/Tmp_ManualTest.bas:2`: `Attribute VB_Name` の直後に `Public Sub` があり、`Option Explicit` / `Option Base 0` / モジュールコメントの順序を満たしていません。
  - 無視する理由: 一時的なファイルであり後で消すため
- [無視] WorkbookService.cls `CommonModules/modules/WorkbookService.cls:12`: `Option Explicit` の直後にモジュールコメントではなく `Implements` が置かれています。
  - 無視する理由: Doxygen 仕様に対応する目的で、正当であるため (コーディング規約の記載を見直したほうがいいかも？)
- [無視] WorksheetService.cls `CommonModules/modules/WorksheetService.cls:12`: `Option Explicit` の直後にモジュールコメントではなく `Implements` が置かれています。
  - 無視する理由: Doxygen 仕様に対応する目的で、正当であるため (コーディング規約の記載を見直したほうがいいかも？)
- [無視] WorksheetServiceTestDouble.cls `CommonModules/modules/WorksheetServiceTestDouble.cls:12`: `Option Explicit` の直後にモジュールコメントではなく `Implements` が置かれています。
  - 無視する理由: Doxygen 仕様に対応する目的で、正当であるため (コーディング規約の記載を見直したほうがいいかも？)

### スコープ未明示のモジュール定数 (14件)

- [要修正] Lib_UnitTest.C_SHEET_NAME (Const) `CommonModules/modules/Lib_UnitTest.bas:11`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_NEW_BOOK (Const) `CommonModules/modules/Lib_UnitTest.bas:12`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COL_MOD (Const) `CommonModules/modules/Lib_UnitTest.bas:13`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COL_SUB (Const) `CommonModules/modules/Lib_UnitTest.bas:14`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COL_OKNG (Const) `CommonModules/modules/Lib_UnitTest.bas:15`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COL_DESC (Const) `CommonModules/modules/Lib_UnitTest.bas:16`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COL_END (Const) `CommonModules/modules/Lib_UnitTest.bas:17`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COL_BTN (Const) `CommonModules/modules/Lib_UnitTest.bas:18`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COL_MBTN (Const) `CommonModules/modules/Lib_UnitTest.bas:19`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_SUB_MAIN (Const) `CommonModules/modules/Lib_UnitTest.bas:20`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COLOR_RESET_BG (Const) `CommonModules/modules/Lib_UnitTest.bas:21`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] Lib_UnitTest.C_COLOR_RESET_FG (Const) `CommonModules/modules/Lib_UnitTest.bas:22`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] ProgressStatus.C_MSG_COMPLETE_BLOCK (Const) `CommonModules/modules/ProgressStatus.cls:20`: モジュール定数のスコープが `Private` / `Public` で明示されていません。
- [要修正] ProgressStatus.C_MSG_UNCOMPLETE_BLOCK (Const) `CommonModules/modules/ProgressStatus.cls:21`: モジュール定数のスコープが `Private` / `Public` で明示されていません。

### Private 補助手続き名 (2件)

- [要修正] WorksheetService.SortCore (Sub) `CommonModules/modules/WorksheetService.cls:113`: Private 補助手続きですが、`p` + PascalCase 名ではありません。
- [要修正] WorksheetServiceTestDouble.SortCore (Sub) `CommonModules/modules/WorksheetServiceTestDouble.cls:173`: Private 補助手続きですが、`p` + PascalCase 名ではありません。

### ローカル変数名 (130件 / 17ファイル)

- [要修正] CounterSet.cls (1件) `CommonModules/modules/CounterSet.cls`: ローカル変数名に `_` を含まないものがあります。73:result
- [要修正] DebugInformation.cls (4件) `CommonModules/modules/DebugInformation.cls`: ローカル変数名に `_` を含まないものがあります。119:idx, 139:result, 143:idx, 168:result
- [要修正] FileSystemService.cls (1件) `CommonModules/modules/FileSystemService.cls`: ローカル変数名に `_` を含まないものがあります。335:idx
- [要修正] Lib_Common.bas (57件) `CommonModules/modules/Lib_Common.bas`: ローカル変数名に `_` を含まないものがあります。327:result, 350:result, 513:shp, 537:idx, 624:result, 652:result, 661:idx, 823:result, 842:result, 843:idx, 862:result, 863:idx, 1142:idx, 1190:result, 1233:result, 1251:idx, 1278:result, 1316:result, 1350:result, 1712:idx, 1722:result, 1754:result, 1770:result, 1771:idx, 1853:result, 1859:idx, 1876:result, 1882:idx, 1929:result, 1930:idx, 1987:char, 1988:idx, 1989:result, 2064:result, 2169:result, 2175:idx, 2194:result, 2200:idx, 2219:result, 2225:idx, 2244:result, 2250:idx, 2275:idx, 2276:result, 2377:result, 2461:result, 2592:result, 2623:result, 2654:result, 2792:result, 2805:result, 2818:result, 2861:result, 2874:result, 2989:result, 3127:result, 3201:result
- [要修正] Lib_IPv4.bas (14件) `CommonModules/modules/Lib_IPv4.bas`: ローカル変数名に `_` を含まないものがあります。66:result, 96:result, 131:result, 194:result, 234:idx, 301:regex, 302:matches, 304:idx, 306:octet, 308:result, 350:result, 352:mask, 426:idx, 465:result
- [要修正] ObjectList.cls (21件) `CommonModules/modules/ObjectList.cls`: ローカル変数名に `_` を含まないものがあります。61:idx, 94:idx, 126:idx, 145:result, 146:idx, 172:result, 173:idx, 224:idx, 258:idx, 272:idx, 287:idx, 307:idx, 352:idx, 410:idx, 505:idx, 524:idx, 548:idx, 564:result, 620:result, 637:result, 643:idx
- [要修正] ObjectSet.cls (9件) `CommonModules/modules/ObjectSet.cls`: ローカル変数名に `_` を含まないものがあります。101:result, 102:idx, 132:result, 133:idx, 203:ItemObject, 218:idx, 233:idx, 247:result, 388:result
- [要修正] ProgressStatus.cls (3件) `CommonModules/modules/ProgressStatus.cls`: ローカル変数名に `_` を含まないものがあります。214:result, 226:result, 228:idx
- [要修正] TextFileEntity.cls (1件) `CommonModules/modules/TextFileEntity.cls`: ローカル変数名に `_` を含まないものがあります。238:result
- [要修正] TextFileService.cls (1件) `CommonModules/modules/TextFileService.cls`: ローカル変数名に `_` を含まないものがあります。22:result
- [要修正] UnitTestAssert.cls (1件) `CommonModules/modules/UnitTestAssert.cls`: ローカル変数名に `_` を含まないものがあります。575:idx
- [要修正] UnitTestUtils.cls (2件) `CommonModules/modules/UnitTestUtils.cls`: ローカル変数名に `_` を含まないものがあります。78:result, 81:idx
- [要修正] UserInputSheet.cls (1件) `CommonModules/modules/UserInputSheet.cls`: ローカル変数名に `_` を含まないものがあります。155:result
- [要修正] WorkbookService.cls (1件) `CommonModules/modules/WorkbookService.cls`: ローカル変数名に `_` を含まないものがあります。434:idx
- [要修正] WorkbookServiceTestDouble.cls (3件) `CommonModules/modules/WorkbookServiceTestDouble.cls`: ローカル変数名に `_` を含まないものがあります。270:result, 316:result, 369:result
- [要修正] WorksheetRangeBounds.cls (5件) `CommonModules/modules/WorksheetRangeBounds.cls`: ローカル変数名に `_` を含まないものがあります。69:result, 81:result, 533:result, 586:result, 757:result
- [要修正] WorksheetService.cls (5件) `CommonModules/modules/WorksheetService.cls`: ローカル変数名に `_` を含まないものがあります。95:idx, 128:idx, 365:result, 455:result, 1211:idx

### Integer 型の index/count 用途 (40件)

- [要修正] Lib_Common.bas `CommonModules/modules/Lib_Common.bas:731`: index / count / length 系の値に `Integer` が使われています。`Dim last_sep As Integer`
- [要修正] Lib_Common.bas `CommonModules/modules/Lib_Common.bas:746`: index / count / length 系の値に `Integer` が使われています。`Dim last_period As Integer`
- [要修正] Lib_Common.bas `CommonModules/modules/Lib_Common.bas:1080`: index / count / length 系の値に `Integer` が使われています。`Dim dim_count As Integer`
- [要修正] Lib_Common.bas `CommonModules/modules/Lib_Common.bas:1985`: index / count / length 系の値に `Integer` が使われています。`ByVal LengthByte As Integer)`
- [要修正] Lib_Common.bas `CommonModules/modules/Lib_Common.bas:2023`: index / count / length 系の値に `Integer` が使われています。`Dim head_pos As Integer`
- [要修正] Lib_Common.bas `CommonModules/modules/Lib_Common.bas:2024`: index / count / length 系の値に `Integer` が使われています。`Dim tail_pos As Integer`
- [要修正] Lib_Common.bas `CommonModules/modules/Lib_Common.bas:2275`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:65`: index / count / length 系の値に `Integer` が使われています。`Public Function ExpandNetwork(ByVal NetworkAddressValue As Long, ByVal MaskLength As Integer) As Long`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:95`: index / count / length 系の値に `Integer` が使われています。`Public Function NarrowNetwork(ByVal NetworkAddressValue As Long, ByVal MaskLength As Integer) As Long()`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:153`: index / count / length 系の値に `Integer` が使われています。`Public Sub ParseIpAddressAndMask(ByRef IpAddressValue As Long, ByRef MaskValue As Long, ByRef MaskLength As Integer, ByVal IpAddressAndMask As String)`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:157`: index / count / length 系の値に `Integer` が使われています。`Dim mask_length As Integer`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:234`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:304`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:405`: index / count / length 系の値に `Integer` が使われています。`Public Function ConvertFromMaskLength(ByVal MaskLength As Integer) As Long`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:425`: index / count / length 系の値に `Integer` が使われています。`Dim mask_length As Integer`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:426`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:464`: index / count / length 系の値に `Integer` が使われています。`Public Function GetMaskValue(ByVal Start As Integer, Optional ByVal Finish As Integer = -1) As Long`
- [要修正] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:468`: index / count / length 系の値に `Integer` が使われています。`Dim bit_length As Integer`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:13`: index / count / length 系の値に `Integer` が使われています。`Private Const C_CHUNK_SIZE As Integer = 16`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:24`: index / count / length 系の値に `Integer` が使われています。`Private pDepth As Integer`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:54`: index / count / length 系の値に `Integer` が使われています。`Dim index_arr() As Integer`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:61`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:119`: index / count / length 系の値に `Integer` が使われています。`Dim index_arr() As Integer`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:126`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:212`: index / count / length 系の値に `Integer` が使われています。`Dim index_arr() As Integer`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:224`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:636`: index / count / length 系の値に `Integer` が使われています。`Private Sub pGetIndexArray(ByRef IndexArray() As Integer, ByRef RemainIndex As Long, ByVal ItemIndex As Long)`
- [要修正] ObjectList.cls `CommonModules/modules/ObjectList.cls:643`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:24`: index / count / length 系の値に `Integer` が使われています。`Private pBlockCount As Integer`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:29`: index / count / length 系の値に `Integer` が使われています。`Private pIntervalPercent As Integer`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:30`: index / count / length 系の値に `Integer` が使われています。`Private pNextUpdate As Integer`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:42`: index / count / length 系の値に `Integer` が使われています。`Public Property Get BlockCount() As Integer`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:194`: index / count / length 系の値に `Integer` が使われています。`Dim prog_percent As Integer`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:213`: index / count / length 系の値に `Integer` が使われています。`Private Function pGetNextUpdate(ByVal ProgressPercent As Integer) As Integer`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:225`: index / count / length 系の値に `Integer` が使われています。`Private Function pGetStatusString(ByVal ProgressRate As Double, ByVal ProgressPercent As Integer) As String`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:227`: index / count / length 系の値に `Integer` が使われています。`Dim comp_blocks As Integer`
- [要修正] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:228`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
- [要修正] UnitTestAssert.cls `CommonModules/modules/UnitTestAssert.cls:560`: index / count / length 系の値に `Integer` が使われています。`Dim expect_dimension As Integer`
- [要修正] UnitTestAssert.cls `CommonModules/modules/UnitTestAssert.cls:563`: index / count / length 系の値に `Integer` が使われています。`Dim actual_dimension As Integer`
- [要修正] UnitTestAssert.cls `CommonModules/modules/UnitTestAssert.cls:575`: index / count / length 系の値に `Integer` が使われています。`Dim idx As Integer`
