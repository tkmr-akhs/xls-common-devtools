# 共通モジュール コーディング規約違反一覧

- [1. 対象](#1-対象)
- [2. 無視するもの / 対象外](#2-無視するもの--対象外)
- [3. 判定メモ](#3-判定メモ)
- [4. サマリ](#4-サマリ)
- [5. 未対応のもの (要対応および要確認)](#5-未対応のもの-要対応および要確認)
  - [5.1. モジュール名長 (0件)](#51-モジュール名長-0件)
  - [5.2. 公開グローバル変数名 (0件)](#52-公開グローバル変数名-0件)
  - [5.3. 公開手続き名 (0件)](#53-公開手続き名-0件)
  - [5.4. 公開引数名 (0件)](#54-公開引数名-0件)
  - [5.5. ファクトリ関数配置 (0件)](#55-ファクトリ関数配置-0件)
  - [5.6. Option Explicit 欠落 (0件)](#56-option-explicit-欠落-0件)
  - [5.7. モジュール冒頭順序 / コメント (0件)](#57-モジュール冒頭順序--コメント-0件)
  - [5.8. スコープ未明示のモジュール定数 (0件)](#58-スコープ未明示のモジュール定数-0件)
  - [5.9. Private 補助手続き名 (0件)](#59-private-補助手続き名-0件)
  - [5.10. ローカル変数名 (0件 / 0ファイル)](#510-ローカル変数名-0件--0ファイル)
  - [5.11. Integer 型の index/count 用途 (0件)](#511-integer-型の-indexcount-用途-0件)
- [6. 保留しているもの](#6-保留しているもの)
  - [6.1. モジュール名長 (3件)](#61-モジュール名長-3件)
  - [6.2. 公開グローバル変数名 (0件)](#62-公開グローバル変数名-0件)
  - [6.3. 公開手続き名 (0件)](#63-公開手続き名-0件)
  - [6.4. 公開引数名 (0件)](#64-公開引数名-0件)
  - [6.5. ファクトリ関数配置 (0件)](#65-ファクトリ関数配置-0件)
  - [6.6. Option Explicit 欠落 (0件)](#66-option-explicit-欠落-0件)
  - [6.7. モジュール冒頭順序 / コメント (0件)](#67-モジュール冒頭順序--コメント-0件)
  - [6.8. スコープ未明示のモジュール定数 (0件)](#68-スコープ未明示のモジュール定数-0件)
  - [6.9. Private 補助手続き名 (0件)](#69-private-補助手続き名-0件)
  - [6.10. ローカル変数名 (0件 / 0ファイル)](#610-ローカル変数名-0件--0ファイル)
  - [6.11. Integer 型の index/count 用途 (0件)](#611-integer-型の-indexcount-用途-0件)
- [7. 対応不要なもの](#7-対応不要なもの)
  - [7.1. モジュール名長 (0件)](#71-モジュール名長-0件)
  - [7.2. 公開グローバル変数名 (2件)](#72-公開グローバル変数名-2件)
  - [7.3. 公開手続き名 (0件)](#73-公開手続き名-0件)
  - [7.4. 公開引数名 (0件)](#74-公開引数名-0件)
  - [7.5. ファクトリ関数配置 (2件)](#75-ファクトリ関数配置-2件)
  - [7.6. Option Explicit 欠落 (0件)](#76-option-explicit-欠落-0件)
  - [7.7. モジュール冒頭順序 / コメント (6件)](#77-モジュール冒頭順序--コメント-6件)
  - [7.8. スコープ未明示のモジュール定数 (0件)](#78-スコープ未明示のモジュール定数-0件)
  - [7.9. Private 補助手続き名 (0件)](#79-private-補助手続き名-0件)
  - [7.10. ローカル変数名 (0件 / 0ファイル)](#710-ローカル変数名-0件--0ファイル)
  - [7.11. Integer 型の index/count 用途 (17件)](#711-integer-型の-indexcount-用途-17件)

## 1. 対象

- 対象ディレクトリ: `CommonModules/modules`
- 対象ファイル: `.bas` / `.cls`
- 除外: `Test_*.bas` / `Test_*.cls`。ユニットテスト本体は今回の洗い出し対象外です。
- `Lib_UnitTest.bas`、`UnitTestAssert.cls`、`UnitTestUtils.cls`、`*TestDouble.cls` は `Test_` で始まらない共通モジュールとして対象に含めています。
- 量があるため、`CommonModules/todo.md` とは別管理です。

## 2. 無視するもの / 対象外

- ユニットテスト本体 (`Test_*`) の規約違反。
- 直接公開フィールドであること自体。改訂後の規約で、読み書き両方を受け入れる場合は外部公開が許容されました。
- シートから利用される大文字英数のみの公開手続き名。
- `ParamArray` 引数に `ByVal` / `ByRef` がないこと。VBA の構文制約として扱います。
- バグ、仕様改善、リファクタリング候補は `CommonModules/todo.md` 側で扱います。

## 3. 判定メモ

- `要修正`: 規約上そのまま違反として扱えるもの。
- `要確認`: 例外や既存仕様として許容される可能性があるため、修正前に意図を確認したいもの。
- `保留`: 規約違反候補だが、互換性や公開 API への影響によりすぐには修正しないもの。
- `対応不要`: 機械的な洗い出しでリストアップされたが、規約違反として扱わないもの。
- ローカル変数名と `Integer` 型の index/count 用途は機械検出を含みます。誤検出の可能性がある項目は `要確認` としています。

## 4. サマリ

| ルール | 要修正 | 要確認 | 保留 | 対応不要 | 合計 |
| --- | ---: | ---: | ---: | ---: | ---: |
| モジュール名長 | 0 | 0 | 3 | 0 | 3 |
| 公開グローバル変数名 | 0 | 0 | 0 | 2 | 2 |
| 公開手続き名 | 0 | 0 | 0 | 0 | 0 |
| 公開引数名 | 0 | 0 | 0 | 0 | 0 |
| ファクトリ関数配置 | 0 | 0 | 0 | 2 | 2 |
| Option Explicit 欠落 | 0 | 0 | 0 | 0 | 0 |
| モジュール冒頭順序 / コメント | 0 | 0 | 0 | 6 | 6 |
| スコープ未明示のモジュール定数 | 0 | 0 | 0 | 0 | 0 |
| Private 補助手続き名 | 0 | 0 | 0 | 0 | 0 |
| ローカル変数名 | 0 | 0 | 0 | 0 | 0 |
| Integer 型の index/count 用途 | 0 | 0 | 0 | 17 | 17 |

## 5. 未対応のもの (要対応および要確認)

ここは、未対応のものが列挙されます。機械的に洗い出され、仕分けされていないものが含まれる可能性があります。

### 5.1. モジュール名長 (0件)

### 5.2. 公開グローバル変数名 (0件)

### 5.3. 公開手続き名 (0件)

### 5.4. 公開引数名 (0件)

### 5.5. ファクトリ関数配置 (0件)

### 5.6. Option Explicit 欠落 (0件)

### 5.7. モジュール冒頭順序 / コメント (0件)

### 5.8. スコープ未明示のモジュール定数 (0件)

### 5.9. Private 補助手続き名 (0件)

### 5.10. ローカル変数名 (0件 / 0ファイル)

### 5.11. Integer 型の index/count 用途 (0件)

## 6. 保留しているもの

ここに残るものは、互換性、公開 API、テスト基盤、または広範囲の機械的リネームに影響するため保留しています。動作に影響しない明示化でも、対象範囲が広いものはこの一覧に記録します。

### 6.1. モジュール名長 (3件)

- [保留] ApplicationScreenUpdateManager (Module) `CommonModules/modules/ApplicationScreenUpdateManager.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=30
  - 保留する理由: 利便性および互換性のため
- [保留] FileSystemServiceTestDouble (Module) `CommonModules/modules/FileSystemServiceTestDouble.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=27
  - 保留する理由: 利便性および互換性のため
- [保留] WorksheetRangeBoundsEnumerator (Module) `CommonModules/modules/WorksheetRangeBoundsEnumerator.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=30
  - 保留する理由: 利便性および互換性のため

### 6.2. 公開グローバル変数名 (0件)

### 6.3. 公開手続き名 (0件)

### 6.4. 公開引数名 (0件)

### 6.5. ファクトリ関数配置 (0件)

### 6.6. Option Explicit 欠落 (0件)

### 6.7. モジュール冒頭順序 / コメント (0件)

### 6.8. スコープ未明示のモジュール定数 (0件)

### 6.9. Private 補助手続き名 (0件)

### 6.10. ローカル変数名 (0件 / 0ファイル)

### 6.11. Integer 型の index/count 用途 (0件)

## 7. 対応不要なもの

### 7.1. モジュール名長 (0件)

### 7.2. 公開グローバル変数名 (2件)

- [対応不要] Lib_FileSystem.FsSrv (Variable) `CommonModules/modules/Lib_FileSystem.bas:13`: 利用者から見える Public 変数です。基盤サービスとして公開名を許容するか確認してください。
  - 対応不要とする理由: 規約上許容するため
- [対応不要] Lib_TextFile.TfSrv (Variable) `CommonModules/modules/Lib_TextFile.bas:13`: 利用者から見える Public 変数です。基盤サービスとして公開名を許容するか確認してください。
  - 対応不要とする理由: 規約上許容するため

### 7.3. 公開手続き名 (0件)

### 7.4. 公開引数名 (0件)

### 7.5. ファクトリ関数配置 (2件)

- [対応不要] New_RangeBounds (Function) `CommonModules/modules/Lib_Common.bas:319`: 引数付きコンストラクター相当の `New_Xxx` 関数ですが、ファクトリ用の `Constructor.bas` ではなく `Lib_Common.bas` にあります。
  - 対応不要とする理由: 規約上許容するため
- [対応不要] New_InputSheet (Function) `CommonModules/modules/Lib_Common.bas:349`: 引数付きコンストラクター相当の `New_Xxx` 関数ですが、ファクトリ用の `Constructor.bas` ではなく `Lib_Common.bas` にあります。
  - 対応不要とする理由: 規約上許容するため

### 7.6. Option Explicit 欠落 (0件)

### 7.7. モジュール冒頭順序 / コメント (6件)

- [対応不要] ArrayObject.cls `CommonModules/modules/ArrayObject.cls:12`: `Option Explicit` の直後にモジュールコメントがなく、Private フィールド宣言が始まっています。
  - 対応不要とする理由: DoxyVB6 のコメント解釈や VBA の記述順の都合として規約上許容するため
- [対応不要] ObjectList.cls `CommonModules/modules/ObjectList.cls:13`: `Option Base 0` の直後にモジュールコメントではなく `Private Const` が置かれています。
  - 対応不要とする理由: DoxyVB6 のコメント解釈や VBA の記述順の都合として規約上許容するため
- [対応不要] Tmp_ManualTest.bas `CommonModules/modules/Tmp_ManualTest.bas:3`: `Option Explicit` の直後に `Public Sub` があり、モジュールコメントがありません。
  - 対応不要とする理由: 一時的なファイルであり後で消すため
- [対応不要] WorkbookService.cls `CommonModules/modules/WorkbookService.cls:12`: `Option Explicit` の直後にモジュールコメントではなく `Implements` が置かれています。
  - 対応不要とする理由: DoxyVB6 のコメント解釈や VBA の記述順の都合として規約上許容するため
- [対応不要] WorksheetService.cls `CommonModules/modules/WorksheetService.cls:12`: `Option Explicit` の直後にモジュールコメントではなく `Implements` が置かれています。
  - 対応不要とする理由: DoxyVB6 のコメント解釈や VBA の記述順の都合として規約上許容するため
- [対応不要] WorksheetServiceTestDouble.cls `CommonModules/modules/WorksheetServiceTestDouble.cls:12`: `Option Explicit` の直後にモジュールコメントではなく `Implements` が置かれています。
  - 対応不要とする理由: DoxyVB6 のコメント解釈や VBA の記述順の都合として規約上許容するため

### 7.8. スコープ未明示のモジュール定数 (0件)

### 7.9. Private 補助手続き名 (0件)

### 7.10. ローカル変数名 (0件 / 0ファイル)

### 7.11. Integer 型の index/count 用途 (17件)

- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:65`: index / count / length 系の値に `Integer` が使われています。`Public Function ExpandNetwork(ByVal NetworkAddressValue As Long, ByVal MaskLength As Integer) As Long`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:95`: index / count / length 系の値に `Integer` が使われています。`Public Function NarrowNetwork(ByVal NetworkAddressValue As Long, ByVal MaskLength As Integer) As Long()`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:153`: index / count / length 系の値に `Integer` が使われています。`Public Sub ParseIpAddressAndMask(ByRef IpAddressValue As Long, ByRef MaskValue As Long, ByRef MaskLength As Integer, ByVal IpAddressAndMask As String)`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:157`: index / count / length 系の値に `Integer` が使われています。`Dim mask_length As Integer`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:234`: index / count / length 系の値に `Integer` が使われています。`Dim item_idx As Integer`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:304`: index / count / length 系の値に `Integer` が使われています。`Dim item_idx As Integer`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:405`: index / count / length 系の値に `Integer` が使われています。`Public Function ConvertFromMaskLength(ByVal MaskLength As Integer) As Long`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:425`: index / count / length 系の値に `Integer` が使われています。`Dim mask_length As Integer`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:426`: index / count / length 系の値に `Integer` が使われています。`Dim item_idx As Integer`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:464`: index / count / length 系の値に `Integer` が使われています。`Public Function GetMaskValue(ByVal Start As Integer, Optional ByVal Finish As Integer = -1) As Long`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] Lib_IPv4.bas `CommonModules/modules/Lib_IPv4.bas:468`: index / count / length 系の値に `Integer` が使われています。`Dim bit_length As Integer`
  - 対応不要とする理由: IPv4 の範囲内であり、Long である必要性がないため
- [対応不要] ObjectList.cls `CommonModules/modules/ObjectList.cls:13`: index / count / length 系の値に `Integer` が使われています。`Private Const C_CHUNK_SIZE As Integer = 16`
  - 対応不要とする理由: インデックス用途ではないため
- [対応不要] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:29`: index / count / length 系の値に `Integer` が使われています。`Private pIntervalPercent As Integer`
  - 対応不要とする理由: インデックス用途ではないため
- [対応不要] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:30`: index / count / length 系の値に `Integer` が使われています。`Private pNextUpdate As Integer`
  - 対応不要とする理由: インデックス用途ではないため
- [対応不要] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:194`: index / count / length 系の値に `Integer` が使われています。`Dim prog_percent As Integer`
  - 対応不要とする理由: インデックス用途ではないため
- [対応不要] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:213`: index / count / length 系の値に `Integer` が使われています。`Private Function pGetNextUpdate(ByVal ProgressPercent As Integer) As Integer`
  - 対応不要とする理由: インデックス用途ではないため
- [対応不要] ProgressStatus.cls `CommonModules/modules/ProgressStatus.cls:225`: index / count / length 系の値に `Integer` が使われています。`Private Function pGetStatusString(ByVal ProgressRate As Double, ByVal ProgressPercent As Integer) As String`
  - 対応不要とする理由: インデックス用途ではないため
