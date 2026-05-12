# 共通モジュール 公開名の規約違反一覧

生成日: 2026-05-12

## 対象

- 対象ディレクトリ: `CommonModules/modules`
- 対象ファイル: `.bas` / `.cls`
- `CommonModules/todo.md` とは別管理です。
- この一覧は「利用者側から見える名称」だけを対象にしています。
- 改訂後の規約により、直接公開フィールドであること自体と、シートから利用される大文字英数の公開手続き名は対象外にしました。

## 含める名称

- モジュール名 / クラス名 / インターフェイス名
- `Public` またはスコープ未明示で公開扱いになる手続き名・プロパティ名
- `Public Const` / `Public` 変数名
- 公開手続きの引数名。VBA では名前付き引数として呼び出し側に見えるため対象にしています。

## 除外するもの

- `Private` メンバー、ローカル変数、Private フィールド
- 直接公開フィールドであること自体。改訂後の規約で、読み書き両方を受け入れる場合は外部公開が許容されました。
- シートから利用される大文字英数のみの公開手続き名
- `Err.Raise`、`On Error`、`Debug.Print`、Excel 直接参照など、名称ではない実装上の規約違反候補
- コメント不足、文字コード、改行など、利用者側から見える名称ではないもの

## 判定メモ

- `要修正`: 規約上そのまま違反として扱えるもの。
- `要確認`: 例外や既存仕様として許容される可能性があるため、修正前に意図を確認したいもの。
- 動作影響なしで明示できるスコープ、戻り値型、引数型、`ByVal`/`ByRef`、Optional 既定値は 2026-05-12 に対応済みです。

## サマリ

| ルール | 要修正 | 要確認 | 合計 |
| --- | ---: | ---: | ---: |
| モジュール名長 | 0 | 12 | 12 |
| 公開グローバル変数名 | 0 | 2 | 2 |

## 残件

動作に影響しない明示化は対応済みです。ここに残るものは、命名・公開状態の整理であり、既存利用者への影響確認が必要なため保留しています。

## 詳細

### モジュール名長 (12件)

- [要確認] ApplicationScreenUpdateManager (Module) `CommonModules/modules/ApplicationScreenUpdateManager.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=30
- [要確認] FileSystemServiceTestDouble (Module) `CommonModules/modules/FileSystemServiceTestDouble.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=27
- [要確認] Test_ApplicationScreenUpdateMan (Module) `CommonModules/modules/Test_ApplicationScreenUpdateMan.bas:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=31
- [要確認] Test_ObjectSetDupCheckDouble (Module) `CommonModules/modules/Test_ObjectSetDupCheckDouble.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=28
- [要確認] Test_ObjectSetEquatableDouble (Module) `CommonModules/modules/Test_ObjectSetEquatableDouble.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=29
- [要確認] Test_ObjectSetKeyPriorityDouble (Module) `CommonModules/modules/Test_ObjectSetKeyPriorityDouble.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=31
- [要確認] Test_TextFileEntityTestDouble (Module) `CommonModules/modules/Test_TextFileEntityTestDouble.bas:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=29
- [要確認] Test_TextFileServiceTestDouble (Module) `CommonModules/modules/Test_TextFileServiceTestDouble.bas:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=30
- [要確認] Test_WorkbookServiceTestDouble (Module) `CommonModules/modules/Test_WorkbookServiceTestDouble.bas:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=30
- [要確認] Test_WorksheetRangeBoundsEnumer (Module) `CommonModules/modules/Test_WorksheetRangeBoundsEnumer.bas:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=31
- [要確認] Test_WorksheetServiceTestDouble (Module) `CommonModules/modules/Test_WorksheetServiceTestDouble.bas:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=31
- [要確認] WorksheetRangeBoundsEnumerator (Module) `CommonModules/modules/WorksheetRangeBoundsEnumerator.cls:1`: Test_ プレフィックスを考慮した推奨 26 文字を超えています。Length=30

### 公開グローバル変数名 (2件)

- [要確認] Lib_FileSystem.FsSrv (Variable) `CommonModules/modules/Lib_FileSystem.bas:13`: 利用者から見える Public 変数です。基盤サービス以外の公開状態か確認してください。
- [要確認] Lib_TextFile.TfSrv (Variable) `CommonModules/modules/Lib_TextFile.bas:13`: 利用者から見える Public 変数です。基盤サービス以外の公開状態か確認してください。
