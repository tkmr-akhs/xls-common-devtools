# 共通モジュール TODO

このファイルは、現時点で見つかっている未対応事項や設計上の検討事項をまとめたものです。

## タグ一覧

| タグ名 | 短縮タグ | 意味 |
| --- | --- | --- |
| バグ | `[bug]` | 現仕様・期待動作に対する誤動作。実務データで誤った結果になるもの。 |
| 要リファクタリング | `[ref]` | 動いてはいるが、設計・責務・依存関係・テスト容易性に問題があるもの。 |
| 仕様改善 | `[spec]` | 仕様の明確化、境界条件の整理、仕様上の矛盾、現仕様として明確だが著しく不便なものの見直し。 |
| 必須機能不足 | `[req]` | 実務利用に必要だが、現時点で仕様として欠けているもの。 |
| 利便性向上 | `[ux]` | 結果の見やすさ、調査しやすさ、操作性、運用しやすさの改善。 |
| 製品試験増強 | `[test]` | 製品試験の考慮不足など。 |

## 高優先度

- [ ] [bug] Lib_UnitTest のテスト候補検出でコメント行を誤実行しない
  - 詳細: `pRunAllTest` の正規表現は行頭の `'` コメントは除外できるが、`Rem Public Sub Test_...(Assert As UnitTestAssert)` のような `Rem` コメントや、コメント化されたテスト宣言風の行を構文として判定し得る。
  - 影響: 無効化したテストや説明コメントがテスト候補として結果シートに出力され、実行時には存在しない手続き呼び出しとして runner error になる。テスト失敗なのかランナーの検出誤りなのかを切り分けにくい。
  - 対応案: VBIDE の手続き情報で列挙するか、少なくともコメント行・属性行を構文として除外したうえで宣言行だけを検出する。`'` コメント、`Rem` コメント、通常の `Public Sub` / `Sub` の検出テストまたは手動プローブを追加する。

- [ ] [bug] Lib_UnitTest の実行用一時モジュールを既存モジュールと衝突させない
  - 詳細: `UnitTestMain` は開始時と終了時に `pRemoveRuntimeRunnerModule` で `UnitTestRuntimeRunner` という名前の VBComponent を無条件に削除し、各テスト実行時にも同名コンポーネントの全コード行を削除してランナーコードを書き込む。
  - 影響: 利用側が同名の通常モジュールを持っていた場合、テストランナーがユーザー定義モジュールを削除または上書きする。ランナーの一時ファイルであることを示すマーカーもなく、衝突時の診断が難しい。
  - 対応案: 一時モジュールに固有マーカーコメントを入れ、削除・上書き前にマーカーと標準モジュール種別を確認する。既存同名モジュールがマーカーなしの場合は `Class Lib_UnitTest` の明示エラーにするテストを追加する。

- [ ] [bug] Lib_UnitTest.UnitTestMain の結果シート AutoFilter を再実行でトグルしない
  - 詳細: `pPrepareResultSheet` は全件実行・単体再実行のどちらでも `Range(...).AutoFilter` を無条件に呼ぶ。既に AutoFilter が設定された結果シートで再実行すると、Excel の AutoFilter が解除または再設定され、利用者のフィルタ状態やドロップダウン表示が変わり得る。
  - 影響: テスト結果を絞り込んだ状態で再実行したときに、結果シートの表示状態がテストランナーの副作用で変わる。単体再実行では、実行結果の更新以外の UI 状態変更が混ざる。
  - 対応案: 結果シート作成時だけ AutoFilter を設定し、既存シートでは `AutoFilterMode` / `FilterMode` を見て必要な範囲更新だけ行う。全件実行と単体再実行で AutoFilter 状態が維持されるテストまたは手動確認を追加する。

- [ ] [bug] UnitTestUtils の配列引数と区切り文字入り同一性キーを安全に扱う
  - 詳細: `UnitTestUtils.pGetKeyCore` は非オブジェクト引数を `pGetPrimitiveKey` へ渡し、最終的に `CStr(ArgItem)` するため、テストダブル対象メソッドの引数を配列のままキー化すると型不一致になる。
  - 詳細: `IEquatable` 引数では `GetIdentityString()` をそのまま `|` 連結に使い、プリミティブ値のような `|` / `<` / `>` のエスケープを行っていない。実装クラスの同一性文字列に区切り文字が入ると、複数引数キーの衝突や誤読が起こり得る。
  - 影響: 配列引数を持つ API を個別変換なしでテストダブルに記録できず、任意文字列を同一性に含む値オブジェクトではスタブ値やスパイ結果の照合が不安定になる。
  - 対応案: 配列引数は要素型・次元・境界を含めてキー化するか、サポート外として `Class UnitTestUtils` の明示エラーにする。`IEquatable` / オブジェクト系キーにも同じエスケープと型タグ規則を適用し、衝突テストを追加する。

- [ ] [bug] UnitTestAssert の通常比較で配列引数を実行時エラーにしない
  - 詳細: `UnitTestAssert.Equals` / `NotEquals` / `EqualsNumeric` / `NotEqualsNumeric` は `pEqualsCore` で配列を特別扱いせず、プリミティブ比較経路の `ExpectedValue = ActualValue` へ進む。`Variant` 配列同士、または配列と非配列を誤って渡すと、アサーション失敗ではなく型不一致の実行時エラーになる。
  - 影響: テストコードで `EqualsArray` を使い忘れた場合に、期待値・実値の診断メッセージを持つ NG ではなく、テストランナー上の ERR として記録される。失敗原因が比較対象の不一致なのか、テスト対象コードの実行時エラーなのかを切り分けにくい。
  - 対応案: `IsArray` を `pEqualsCore` の入口で検出し、両方配列なら `pEqualsArrayCore` へ委譲するか、通常比較では配列非対応として `UnitTestAssert` の明示的な失敗メッセージにする。配列同士、配列とスカラー、数値型無視比較へ配列を渡したケースのテストを追加する。

- [ ] [bug] UnitTestAssert の数値比較で非数値プリミティブを実行時エラーにしない
  - 詳細: `EqualsNumeric` / `NotEqualsNumeric` は `pEqualsPrimitiveNumericCore` で `ExpectedValue = ActualValue` を実行してから `pIsNumericType` を確認している。`"abc"` と `1` のように暗黙変換できない非数値プリミティブを渡すと、アサーション失敗ではなく型不一致の実行時エラーになる。
  - 影響: 数値比較アサーションの使い方を誤ったテストが `NG` ではなく `ERR` になり、テスト対象コードの実行時エラーとアサーション入力の不備を切り分けにくい。
  - 対応案: 数値型・数値文字列として扱う範囲を比較前に判定し、対象外は明示的なアサーション不一致として記録する。非数値文字列、数値文字列、Date、Currency / Decimal の扱いをテストで固定する。

- [ ] [bug] Lib_UnitTest の結果シート未作成経路で内部 Err.Number を残さない
  - 詳細: `pPrepareResultSheet` は全件実行時に `On Error Resume Next` で `ThisWorkbook.Worksheets(C_SHEET_NAME)` を探すが、結果シートがない通常初回実行で発生したエラーを `Err.Clear` しないまま新規シート作成へ進む。
  - 影響: `UnitTestMain` 自体は成功しても、呼び出し側や自動化側が `On Error Resume Next` 後の `Err.Number` を確認する運用では、正常な初回実行を失敗扱いする可能性がある。テストランナー内部の探索エラーと実際のテスト失敗も切り分けにくい。
  - 対応案: 想定内の未検出エラーは分岐直後に `Err.Clear` し、結果シート既存/未作成それぞれで `Err.Number = 0` のまま戻ることを手動プローブまたはテストで固定する。

- [ ] [bug] WorksheetRangeBounds.Item の負インデックスを範囲外として拒否する
  - 詳細: `WorksheetRangeBounds.Item` は `pGetCellVertical` / `pGetCellHorizontal` で `ItemIndex Mod ...` から行・列を計算するが、`ItemIndex < 0` を入口で拒否していない。負数を渡すと、範囲外エラーではなく開始位置より前のセルを返したり、正規化された別範囲になる可能性がある。
  - 影響: `WorksheetRangeBoundsEnumerator` 経由では通常発生しないが、公開 API として `Item(-1)` が意図しないセルを返すと、呼び出し側が範囲外アクセスを検出できず誤ったセルを読み書きする。
  - 対応案: `Item` の入口で `ItemIndex < 0 Or Count <= ItemIndex` を範囲外として明示エラーにする。縦方向・横方向それぞれで `-1`、`0`、`Count - 1`、`Count` のテストを追加する。

- [ ] [bug] WorksheetRangeBounds.ToString の絶対 A1 アドレス生成を失敗させない
  - 詳細: `WorksheetRangeBounds.ToString(IsAbsoluteStartRow:=True...)` は `RangeAddress` に `ReferenceRow:=0` / `ReferenceColumn:=0` を渡すが、`RangeAddress` は絶対参照指定時に基準行・列が 1 未満だとエラーにしている。絶対参照では基準値を使わないため、`$A$1` のようなアドレス生成が不要に失敗し得る。
  - 詳細: 列範囲の A1 生成では、終了列の絶対指定に `IsAbsoluteFinishColumn` ではなく `IsAbsoluteStartColumn` を使っており、開始列と終了列で絶対/相対指定を変えた場合に表記が崩れる。
  - 影響: `WorksheetRangeBounds` の表示、同一性文字列、診断メッセージ、アドレス生成ヘルパーを絶対参照付きで使う呼び出しが、範囲値自体は正しいのに実行時エラーまたは誤ったアドレスになる。
  - 対応案: A1 の基準値検証は相対参照を生成する場合だけ行い、列範囲の終了列には `IsAbsoluteFinishColumn` を渡す。絶対セル、絶対行範囲、絶対列範囲、開始/終了で絶対指定が異なる列範囲のテストを追加する。

- [ ] [bug] WorksheetService.WriteCell の内部型変換失敗で Err.Number を残さない
  - 詳細: `WriteCell(TypeConvert:=True)` は `CBool` / `CDate` / `CCur` を `On Error Resume Next` で順に試し、すべて失敗した場合に `Err.Clear` しないまま文字列書き込みへ進む。処理自体は成功しても、呼び出し元の `Err.Number` に内部変換失敗の値が残り得る。
  - 影響: 呼び出し側が `On Error Resume Next` と `Err.Number` で `WriteCell` の成否を見ている場合、正常に文字列として書き込めた値を失敗扱いする。テストでも基盤 API 内部の探索的変換エラーと実際の書き込みエラーを区別しづらい。
  - 対応案: 型変換試行の最後で `Err.Clear` してから通常のエラーハンドリングへ戻す。数値/Boolean/Date/Currency に変換できない文字列を書き込んだあと、`Err.Number = 0` で戻るテストを追加する。

- [ ] [bug] WorksheetService.WriteCell の文字列書き込みで表示形式を必ず復元する
  - 詳細: `WriteCell(TypeConvert:=False)` または型変換に失敗した文字列書き込みでは、現在の `NumberFormatLocal` を退避してから一時的に `@` に変更し、値設定後に元へ戻している。値設定や復元前の処理でエラーになると、セルの表示形式が `@` のまま残る。
  - 影響: 書き込み自体は失敗しても、シート側には表示形式変更だけが副作用として残り、後続の数値・日付入力が文字列扱いになる可能性がある。基盤書き込み API として失敗時の状態が予測しにくい。
  - 対応案: 表示形式の変更から復元までをエラーハンドリングで囲み、成功時・失敗時とも元の表示形式へ戻してから再送出する。入力規則違反や保護セルなど、値設定失敗時の復元テストを追加する。

- [ ] [bug] WorksheetService.WriteRange の配列サイズを対象範囲と照合する
  - 詳細: `WriteRange` は `ValuesArray` をそのまま `target_range.Value` に代入しており、配列の次元数、下限、行数、列数が `RangeBounds` と一致するかを確認していない。Excel 側の代入仕様により、不足要素が `#N/A` になったり、余剰要素が書き込まれない可能性がある。
  - 影響: 呼び出し側の配列整形ミスが明示エラーではなくシート上のデータ破壊として表面化する。表データ書き込みの基盤 API として、意図しない `#N/A` や欠落に気付きにくい。
  - 対応案: `ValuesArray` は 2 次元配列かつ `RowCount` / `ColumnCount` と一致することを入口で検証する。1 セル、1 行、1 列、下限 0/1、過不足配列、スカラー入力の契約をテストで固定する。

- [ ] [bug] WorksheetService.InsertRows / DeleteRows で行範囲以外を拒否する
  - 詳細: `InsertRows` / `DeleteRows` は `RangeBounds.Row` から `RangeBounds.FinishRow` だけを使って `"行:行"` の Range を作るため、列範囲や通常のセル範囲を渡しても行全体の挿入・削除として実行される。特に `A:A` のような列範囲では全行相当を対象にし得る。
  - 影響: 呼び出し側が範囲オブジェクトを誤って渡したとき、明示エラーではなくシート構造を大きく変更する可能性がある。基盤 API として、範囲単位操作と行単位操作の境界が危険に曖昧になる。
  - 対応案: `RangeBounds.IsEntireRow` または行方向の範囲だけを受け付ける契約に固定し、列範囲、セル範囲、複数行範囲、空範囲のテストを追加する。

- [ ] [bug] ObjectList / ObjectSet の検索・削除 API で要素型を検査する
  - 詳細: `ObjectList.Exists` / `GetIndexByItem` / `RemoveItem` は特殊値チェックだけで `pCheckItemType` を通さず、プリミティブ比較では `ItemObject1 = ItemObject2` の暗黙変換に任せている。そのため数値 `1` のリストに文字列 `"1"` を渡すと同一扱いになり得る。
  - 詳細: `ObjectSet.Exists` / `GetContains` / `RemoveItem` も追加時と同じ型検査を通さず、`IDuplicateCheckable` / `IEquatable` のキー生成経路では型違いの値が実装詳細由来の実行時エラーになる可能性がある。
  - 影響: 追加・更新では型を固定しているのに、検索・削除では型違いを False、暗黙一致、実行時エラーのどれにするかが揃わない。コレクション基盤として「同じ要素」の意味が操作ごとに変わる。
  - 対応案: 検索・削除系 API でも保存済み型との照合を行い、型違いは明示エラーまたは常に不一致のどちらかに統一する。数値と数値文字列、Boolean と数値、IEquatable 実装と非オブジェクト引数のテストを追加する。

- [ ] [bug] WorksheetService.CopyRange でコピー先がシート端を超える場合を明示的に扱う
  - 詳細: `CopyRange` はコピー先範囲を `New_RangeBounds` で作るため、`FinishRow` / `FinishColumn` が Excel 上限を超えると範囲オブジェクト側で上限へ丸められる。一方でコピー処理のループ回数は `src_bounds.RowCount` / `ColumnCount` と `dst_bounds.RowCount` / `ColumnCount` の大きい方で決まるため、コピー元の方が大きい場合に `DestinationSheet.Cells(DestinationRow + row_idx, ...)` がシート外を参照し得る。
  - 影響: シート最終行・最終列付近へ範囲コピーしたとき、範囲作成時には失敗しないのにコピー途中で Excel の実行時エラーになる。部分コピー、明示エラー、事前拒否のどれが仕様かも呼び出し側から判断しにくい。
  - 対応案: コピー先がシート上限を超える入力を事前に検出して明示エラーにするか、コピー可能な交差範囲だけを処理する仕様に固定する。最終行・最終列付近への `CopyRange` テストを追加する。

- [ ] [bug] WorksheetService.SetSheetOutlineLevel の既定値 0 を Excel に渡さない
  - 詳細: `SetSheetOutlineLevel` は `RowLevels:=0`、`ColumnLevels:=0` を既定値にし、そのまま `target_sheet.Outline.ShowLevels(RowLevels:=RowLevels, ColumnLevels:=ColumnLevels)` へ渡している。Excel のアウトラインレベルは 1 以上の指定または引数省略が前提で、0 を渡すと実行時エラーになり得る。
  - 影響: 引数省略で「変更しない」つもりの呼び出しや、行だけ・列だけを指定する呼び出しが、アウトライン操作ではなく基盤 API 側の既定値で失敗する。
  - 対応案: `0` は未指定を表す値として扱い、行・列それぞれ 1 以上のときだけ `ShowLevels` へ渡す。`0`、行のみ、列のみ、両方指定、範囲外値のテストを追加する。

- [ ] [bug] WorksheetService.Sort のキー列番号と並び順を事前検証する
  - 詳細: `WorksheetService.Sort` は `SortKeyAndOrder` の個数が偶数であることだけを確認し、キー列番号が `RangeBounds` の列数内か、並び順が `xlAscending` / `xlDescending` かを検証していない。`0`、負数、範囲外列、無効な並び順が Excel 呼び出しへそのまま渡る。
  - 影響: 呼び出し側の指定ミスで、対象範囲外の列をキーにしたソートや Excel 由来の実行時エラーが発生する。データ並べ替えの基盤 API として、誤ったデータ移動の原因を特定しにくい。
  - 対応案: キー列番号は `1 To RangeBounds.ColumnCount`、並び順は未指定時の既定値または Excel の有効値だけに制限する。`0`、負数、列数超過、無効な並び順のテストを追加する。

- [ ] [bug] WorksheetService.Sort で空範囲を他の範囲 API と同じく no-op にする
  - 詳細: `WorksheetService.Sort` は `WriteCell`、`WriteRange`、`CopyCell`、`ClearRange` などと異なり、`RangeBounds.IsEmpty` を入口で確認しない。空範囲を渡すと `RangeBounds.ToString(CellOnly:=True)` の空範囲表記を Excel `Range` として解釈しようとして、ソート対象なしの no-op ではなく実行時エラーになり得る。
  - 影響: `GetUsedRangeBounds` や `Intersect` の結果をそのままソートへ渡す呼び出し側が、空データのときだけ個別分岐を強いられる。空範囲を no-op とする他の `WorksheetService` API と契約が揃わない。
  - 対応案: `pSortCore` の先頭で空範囲を検出して終了する。空範囲、キーなし、キーあり空範囲のテストを追加し、空データ処理で例外が出ないことを固定する。

- [ ] [bug] WorksheetService.Find で空範囲を空結果として扱う
  - 詳細: `Find` は `RangeBounds.IsEmpty` を入口で確認せず、空範囲の `ToString(CellOnly:=True)` を Excel `Range` に渡す。`GetUsedRangeBounds` や `Intersect` の結果が空の場合、検索結果 0 件ではなく実行時エラーになり得る。
  - 影響: 空データのシートや空の交差範囲を検索する呼び出し側が、`Find` の前だけ個別分岐を強いられる。`WorkbookService.Find` や `UserInputSheet.GetItemRange` など、検索を基盤 API として使う処理でも空範囲時に落ちる可能性がある。
  - 対応案: `Find` の先頭で空範囲を検出して未初期化配列を返すか、空の `WorksheetRangeBounds()` を返す契約に固定する。空範囲、空シート、非空範囲 0 件のテストを追加する。

- [ ] [bug] WorksheetService.RemoveDuplicates の重複判定列を対象範囲内に検証する
  - 詳細: `RemoveDuplicates` は `DuplicateColumns` をそのまま Excel `Range.RemoveDuplicates` に渡しており、0、負数、対象範囲の列数超過、空配列、絶対列番号の誤指定を基盤側で検出しない。
  - 影響: 呼び出し側の指定ミスが Excel 由来の実行時エラーになるか、意図と違う列を基準にした行削除として実行される可能性がある。行削除系 API として失敗時の影響が大きい。
  - 対応案: 判定列は `1 To RangeBounds.ColumnCount` の相対列番号に統一し、Scalar / Array の双方で全要素を検証する。空配列、0、負数、範囲外、正常複数列のテストを追加する。

- [ ] [bug] WorkbookService.ActivateWorksheet で対象ブックを先にアクティブ化する
  - 詳細: `WorkbookService.ActivateWorksheet` は `target_sheet.Activate` を直接呼ぶだけで、対象ブックを先にアクティブにしていない。
  - 影響: 対象シートがアクティブでない別ブックにある場合、`Worksheet.Activate` が失敗したり、呼び出し側が期待したブック・シートへ移動できない可能性がある。
  - 対応案: UI 操作用 API として対象ブック、対象シートのアクティブ化順序を明示し、非アクティブブック上のシートを指定したテストを追加する。

- [ ] [bug] WorkbookService.CloseWorkbook の確認表示契約を守る
  - 詳細: `WorkbookService.CloseWorkbook` は `Force:=False` のとき確認画面を表示する契約だが、実装は `Workbooks(Book).Close(SaveChanges:=Not Force)` となっており、確認なしで保存して閉じる可能性がある。
  - 影響: 利用者が確認してから保存可否を判断する前提の処理で、意図しない変更をブックへ保存してしまう可能性がある。
  - 対応案: `Force:=False` では `SaveChanges` を省略して Excel の確認に委ねる、または確認を独自に行う。`Force:=True` の保存/破棄方針もインターフェイスコメントと合わせてテストで固定する。

- [ ] [bug] WorkbookService の存在確認・重複名チェックで内部探索エラーを残さない
  - 詳細: `ExistsWorkbook` / `ExistsWorksheet` / `AddWorksheet` / `CopyWorksheet` は、存在しないブックやシートを `On Error Resume Next` で探索するが、想定内の未検出エラーを `Err.Clear` しない経路がある。
  - 影響: API の戻り値や処理自体は成功しても、呼び出し側が `Err.Number` を確認する運用では、存在しないことを調べただけの内部エラーを実処理の失敗として扱う可能性がある。
  - 対応案: 期待される未検出エラーは分岐直後に必ず `Err.Clear` する。存在しないブック、存在しないシート、追加先名未使用、複製先名未使用の各ケースで `Err.Number = 0` を確認するテストを追加する。

- [ ] [bug] WorkbookService.AddWorksheet でアクティブブック・シートを復元する
  - 詳細: `AddWorksheet` は対象ブックへ `Worksheets.Add` するため、Excel のアクティブブックとアクティブシートが追加先・追加シートへ移るが、`CopyWorksheet` や `SaveWorkbook` のような退避復元を行っていない。
  - 影響: 利用側が `WorksheetRangeBounds` や明示ブック名で操作していても、シート追加後の後続処理で `ActiveWorkbook` / `ActiveSheet` に依存する既存コードが別ブック・別シートを対象にしてしまう可能性がある。基盤サービス呼び出しの副作用として UI 状態も変わる。
  - 対応案: 追加前の `ActiveWorkbook` と対象ブックのアクティブシートを退避し、成功時・エラー時とも復元する。非アクティブブックへの追加、ThisWorkbook への追加、既存名エラー時の復元テストを追加する。

- [ ] [bug] WorkbookService.AddWorksheet / CopyWorksheet の名前設定失敗時に追加済みシートを残さない
  - 詳細: `AddWorksheet` は `Worksheets.Add` 後に `added_sheet.Name = Sheet` を実行し、`CopyWorksheet` も `src_sheet.Copy` 後に `added_sheet.Name = DestinationWorksheetName` を実行する。31 文字超過や禁止文字などで名前変更が失敗すると、追加または複製されたシートだけが残る。
  - 影響: API はエラーで戻るのにブック構成は変更済みとなり、リトライ時の重複シート、後続処理の対象ずれ、手作業での残骸削除が発生し得る。
  - 対応案: シート名を追加前に検証するか、名前設定失敗時に追加済みシートを削除してから再送出する。禁止文字、31 文字超過、既存名、正常追加・複製のテストを追加する。

- [ ] [bug] WorkbookService.RemoveWorksheet でアクティブブック・シートを復元する
  - 詳細: `RemoveWorksheet` は削除対象ブックを操作し、最後の 1 シートを削除する場合は代替シートも追加するが、処理前の `ActiveWorkbook` や対象ブックのアクティブシートを退避復元していない。
  - 影響: 非アクティブブックのシート削除や最後のシート削除の後、Excel のアクティブ状態が削除先ブック・追加シートへ移り、後続処理が `ActiveWorkbook` / `ActiveSheet` 依存の場合に対象を誤る可能性がある。
  - 対応案: `CopyWorksheet` と同様に現在のアクティブブックと対象ブックのアクティブシートを退避し、成功時・エラー時とも復元する。最後の 1 シート削除、非アクティブブック削除、削除失敗時の復元テストを追加する。

- [ ] [bug] WorkbookService.RemoveVBComponents の削除前検証を行う
  - 詳細: `RemoveVBComponents` は `VBComponents` を走査しながら一致したモジュールを即時削除するため、後続で一致した `ThisWorkbook` やシートモジュールなど削除できないコンポーネントに当たると、それ以前の標準モジュールだけが削除済みの中途半端な状態になり得る。
  - 詳細: `ComponentNames` は `Variant` だが、`pContainsComponentName` は配列要素やスカラー値を直接 `CStr` するため、`Null`、`CVErr(...)`、未初期化配列などの入力で、削除対象名の検証前に実行時エラーになり得る。
  - 影響: 不正な削除対象指定や削除不能モジュール混在時に、VBProject の一部だけが変更される。モジュール適用前のクリーンアップ API として、失敗時の復旧や診断が難しい。
  - 対応案: 削除前に `ComponentNames` を型付きの名前集合へ正規化し、存在有無、コンポーネント種別、削除可能性を全件検証してから削除する。標準モジュールと Document モジュール混在、`Null` / `CVErr` / 空配列 / 未初期化配列のテストを追加する。

- [ ] [bug] WorksheetService の Excel エラー値文字列化を安全にする
  - 詳細: `WorksheetService.pConvertErrorToString` は `Select Case ErrValue` と `Case CVErr(...)` で Excel エラー値を判定しており、`CVErr(...)` を含む Variant の直接比較で型不一致になる可能性がある。
  - 詳細: `ReadCell(GetText:=True)` の `####` フォールバックでも `"" & TargetCell.Value` を使うため、対象セルがエラー値の場合に文字列化で落ち得る。
  - 影響: `ReadCell`、`XLookup`、`pGetFormulaLiteral` など、Excel エラー値を読み取る・式へ埋め込む経路で、テスト対象ではなく基盤側の変換処理が失敗する。
  - 対応案: `IsError` 分岐後は `WorksheetFunction.Error_Type` 相当の安定した判定に寄せ、7 種類の Excel エラー値を文字列へ変換するテストを追加する。

- [ ] [bug] WorksheetService の通常数式コピーで相対参照を保持する
  - 詳細: `WorksheetService.CopyCell` / `CopyRange` の通常数式コピー分岐で `dst_cell = src_cell.FormulaR1C1` としており、`dst_cell.FormulaR1C1 = src_cell.FormulaR1C1` になっていない。
  - 影響: `=A2` のような相対参照を含む通常数式をコピーしたとき、コピー先で参照先がずれる、または数式として解釈できない値になる可能性がある。
  - 対応案: 通常数式のコピー先プロパティを明示し、A1 形式と R1C1 形式の期待動作を固定する。相対参照を含む通常数式の `CopyCell` / `CopyRange` テストを追加する。

- [ ] [bug] WorksheetService.ReadCell(GetText:=True) の副作用と表示形式戻り値を修正する
  - 詳細: `ReadCell(GetText:=True)` は `Range.Text` を読むために列幅を変更するが、元の `ColumnWidth` を `Long` に保存しているため、小数を含む列幅を正確に復元できない。途中でエラーが発生した場合に列幅が戻らない経路もある。
  - 詳細: `GetText:=False` では `NumberFormatLocal` を返す一方、`GetText:=True` では `NumberFormat` を返しており、インターフェイスコメントと `WriteCell` の表示形式指定と揃っていない。
  - 影響: 読み取り API の呼び出しだけでシートの列幅が変わる可能性があり、表示形式の戻り値もロケール依存の読み書きで不整合になる。
  - 対応案: 列幅は `Double` で退避し、エラー時も復元する。`NumberFormat` の返却値は `NumberFormatLocal` に統一し、列幅保持とローカル表示形式のテストを追加する。

- [ ] [bug] WorksheetService.GetUsedRangeBounds の Find 書式条件依存をなくす
  - 詳細: `pGetLastRowCore` / `pGetLastColumnCore` / `pGetFirstRowCore` / `pGetFirstColumnCore` は `Range.Find(What:="*")` で `SearchFormat:=False` を明示していない。Excel の検索ダイアログなどで書式検索条件が残っていると、値があるセルでも UsedRange の先頭・末尾検出から漏れる可能性がある。
  - 影響: `GetUsedRangeBounds`、`CopyRange`、`UserInputSheet.GetItemRange` など使用範囲に依存する処理が、利用者の直前の検索状態で空範囲や狭い範囲を返し得る。
  - 対応案: UsedRange 検出用の全 `Find` に `SearchFormat:=False` を明示し、検索書式条件が残っている状態で先頭/末尾行列を正しく検出するテストを追加する。

- [ ] [bug] WorksheetService.Find の Excel Find 状態依存をなくす
  - 詳細: `Range.Find` で `After`、`SearchOrder`、`SearchDirection`、`SearchFormat` を明示していない。Excel の検索ダイアログや直前の `Find` 設定が残っていると、同じ引数でも結果が変わり得る。
  - 詳細: 検索設定を戻すためのダミー `Find` は検索結果が 0 件のとき実行されず、成功時も `SearchFormat` を明示的に解除していない。
  - 影響: `WorksheetService.Find` とそれを使う `WorkbookService.Find`、`UserInputSheet.GetItemRange` が、利用者の直前の検索条件や Excel UI 状態に依存して見つかる/見つからないが変わる可能性がある。
  - 対応案: `SearchFormat:=False` を含めて検索条件を明示し、可能ならダミー検索ではなく呼び出し内で完結する設定にする。検索フォーマット条件が残っている状態、0 件、複数件のテストを追加する。

- [ ] [bug] WorksheetService.ActivateRange で対象シートを確実にアクティブ化する
  - 詳細: `ActivateRange` は `target_sheet.Range(...).Activate` を直接呼ぶだけで、対象ブックや対象シートを先にアクティブにしていない。
  - 影響: 対象シートがアクティブでない状態では `Range.Activate` が失敗したり、呼び出し側が期待した範囲が選択されない可能性がある。
  - 対応案: UI 操作用 API として対象ブック、対象シート、対象範囲のアクティブ化順序を明示し、非アクティブブック/非アクティブシートからの呼び出しをテストする。

- [ ] [bug] WorksheetService の配列数式コピーでコピー元を破壊しない
  - 詳細: `WorksheetService.pCopyCellCore` は配列数式コピー時に `src_cell = formula_str` でコピー元セルを通常数式へ一時変更し、後続処理が成功した場合だけ `src_cell.FormulaArray = formula_str` で戻している。
  - 影響: `Application.ConvertFormula`、コピー先への代入、配列数式の再設定でエラーが発生すると、コピー処理の失敗だけでコピー元の配列数式が壊れる可能性がある。読み取り・コピー系 API として副作用が大きい。
  - 対応案: コピー元セルを書き換えずに相対参照を変換する方法へ変更する。やむを得ず一時変更する場合はエラー時も必ず復元するガードを置き、コピー先失敗時にコピー元が維持されるテストを追加する。

- [ ] [bug] WorksheetRangeBounds の未初期化状態を全公開メンバーで拒否する
  - 詳細: `WorksheetRangeBounds.TransformAbsolute` は `pCheckInit` を呼ばず、未初期化インスタンスでも既定値から新しい範囲を返し得る。`GetIdentityString` / `Equals` も `ToString` の `UNINITIALIZED(...)` 表記に寄るため、未初期化状態をキーや比較に流せてしまう。
  - 影響: 呼び出し側の初期化漏れが早期に失敗せず、意図しない `Sheet1` / `ThisWorkbook` 起点の範囲や不安定な同一性文字列として扱われる可能性がある。
  - 対応案: `ToString` の診断用途を残すかを決めたうえで、通常利用の公開メンバーは `pCheckInit` に統一する。未初期化 `TransformAbsolute`、`Equals`、`GetIdentityString` のテストを追加する。

- [ ] [bug] WorksheetRangeBounds.Initialize で開始行・開始列の Excel 上限超過を拒否する
  - 詳細: `pWellFormBounds` は `FinishRow` / `FinishColumn` が Excel 上限を超えた場合は丸めるが、`Row` / `Column` 自体が `G_ROW_MAX` / `G_COL_MAX` を超えた場合はそのまま保持する。開始行が上限超過で終了行だけ丸められると、Excel シート上に存在しない開始位置を持つ空範囲が作れてしまう。
  - 影響: `WorksheetService` 側で `Cells(RangeBounds.Row, RangeBounds.Column)` や `Range(...)` に渡した段階で Excel 実行時エラーになり、範囲値オブジェクトの生成時点で入力不正を検出できない。`TransformAbsolute` やコピー先範囲計算でも同じ不正範囲が伝播する。
  - 対応案: 開始行・開始列が 1 未満または Excel 上限超過の場合は明示エラーにするか、空範囲として許す条件を仕様化する。開始行・開始列・終了行・終了列それぞれの上限超過テストを追加する。

- [ ] [bug] WorksheetRangeBounds.Shift / Transform の加減算桁あふれを事前検出する
  - 詳細: `Shift` は `pRow + Row` / `pFinishRow + Row`、`Transform` は `pFinishRow + AddRow` / `pFinishColumn + AddColumn` を先に `Long` 同士で計算し、その後に 1 未満や Excel 上限超過を検証している。大きな正負の移動量を渡すと、検証前に VBA の Overflow が発生する。
  - 影響: 呼び出し側には `Class WorksheetRangeBounds` の範囲エラーではなく実行時エラー 6 が返り、どの範囲操作が不正だったかを診断しづらい。範囲計算の入力検証としても `Initialize` / `Shift` の契約が揃わない。
  - 対応案: 加減算前に `Long` の範囲と Excel 行列上限を事前判定するか、`Double` など一時的に広い型で計算してから明示エラーにする。上限付近の正方向・負方向シフト、空範囲の `Transform` のテストを追加する。

- [ ] [bug] WorksheetRangeBoundsEnumerator.Initialize で Nothing の列挙対象を拒否する
  - 詳細: `Initialize(TargetCollection As WorksheetRangeBounds)` は `TargetCollection Is Nothing` を確認せず、`Target` に `Nothing` を設定したまま初期化済みにできる。`pCalculateLength` は `pRangeBounds Is Nothing` で何もせず、列挙子は長さ 0 のように振る舞う。
  - 影響: 呼び出し側の範囲生成漏れが初期化時に検出されず、`MoveNext` が False を返すだけの空列挙として扱われる。`Target` や `Current` で後から落ちる場合もあり、基盤列挙子として原因を追いにくい。
  - 対応案: `Initialize` の入口で `Nothing` を `Class WorksheetRangeBoundsEnumerator` の明示エラーにする。`Nothing`、空範囲、通常範囲の初期化テストを追加する。

- [ ] [bug] WorksheetRangeBounds.Intersect で Nothing を明示的に拒否する
  - 詳細: `Intersect(ByVal OtherRangeBounds As WorksheetRangeBounds)` は `OtherRangeBounds Is Nothing` を確認せず、すぐに `OtherRangeBounds.WorksheetName` / `WorkbookName` / `Row` などを参照する。
  - 影響: 呼び出し側の範囲生成漏れが `Class WorksheetRangeBounds` の説明付きエラーではなく、VBA の汎用的な「オブジェクト変数または With ブロック変数が設定されていません。」として表面化する。範囲交差を使う `WorksheetService.GetUsedRangeBounds` などでも原因を追いにくい。
  - 対応案: 入口で `Nothing` を明示エラーにし、`Nothing`、未初期化範囲、空範囲、非交差範囲、正常交差範囲のテストを分けて追加する。

- [ ] [bug] ObjectList / ObjectSet の特殊 Variant 値の扱いを固定する
  - 詳細: `ObjectList.pItemsEqual` はプリミティブ比較で `ItemObject1 = ItemObject2` を直接 Boolean に代入するため、`Null` や `CVErr(...)` で比較自体が失敗し得る。`ObjectSet` は `Null` / `CVErr(...)` / `Empty` を Dictionary キーとして渡す経路があり、`ConvertToStringArray` でも特殊値の文字列化方針がない。
  - 影響: Excel Range から読んだ値や外部入力をそのまま基盤コレクションへ入れると、追加、存在判定、削除、重複除去、文字列化のどこで失敗するかが値によって変わる。
  - 対応案: `Null`、`Empty`、`CVErr(...)` をサポートするか明示エラーにするかを決め、`ObjectList` / `ObjectSet` / `Enumerator` の取得・比較・文字列化テストを追加する。

- [ ] [bug] ObjectList / ObjectSet の配列要素サポート有無を実装と揃える
  - 詳細: `ObjectList.pCheckUnsupportedSpecialValue` と `ObjectSet.pCheckUnsupportedSpecialValue` は配列を許可しているが、`ObjectList.pItemsEqual` は配列同士を `ItemObject1 = ItemObject2` で比較し、`ConvertToStringArray` でも配列を `String` 要素へ直接代入するため型不一致になり得る。`ObjectSet.pGetKey` も配列をそのまま Dictionary キーに渡す経路がある。
  - 影響: 配列要素を追加できるように見える一方で、存在判定、削除、重複削除、文字列化、集合追加のどこかで実行時エラーになる。Range 値や分割結果を要素として扱う利用者が、どの操作まで安全か判断できない。
  - 対応案: 配列要素をサポート外として追加時に明示エラーにするか、次元・境界・要素を含む比較/キー化/文字列化仕様を実装する。`ObjectList` と `ObjectSet` の配列追加、Exists、RemoveItem、RemoveDuplicate、ConvertToStringArray、Add のテストを追加する。

- [ ] [bug] ObjectSet.Update のキー不一致エラーでオブジェクトキーを安全に表示する
  - 詳細: `Update` は旧キーと新キーが異なる場合に `old_key` / `new_key` をそのまま文字列連結してエラー説明を作る。参照同一性をキーにするオブジェクト集合で別インスタンスへ更新しようとすると、意図した `Class ObjectSet` のキー不一致エラーではなく、オブジェクトの文字列変換エラーになり得る。
  - 影響: 更新禁止条件そのものは検出しているのに、呼び出し側には原因が読めるエラーが返らない。`Enumerator.Update` 経由で `ObjectSet` を更新する場合も診断しづらい。
  - 対応案: キー表示用の安全な文字列化ヘルパーを追加し、オブジェクトは `TypeName` と `ObjPtr`、`Nothing`、`CVErr` などを型付きで表示する。オブジェクト参照キー、`IEquatable` キー、`CVErr` キーの不一致テストを追加する。

- [ ] [bug] ObjectList.Sort の等価要素と降順比較を修正する
  - 詳細: `ObjectList.Sort` は `IComparable` でも降順時に `Not IsLessThan` を使うため、等価要素を「小さい」扱いし得る。
  - 影響: 等価要素を含む並び替えで不要な入れ替えが発生し、安定性や比較契約に依存する処理の前提が崩れる可能性がある。
  - 対応案: 昇順・降順とも「小さい」「大きい」「等価」を分けて判定し、等価要素を入れ替えないテストを追加する。

- [ ] [bug] ObjectList / ObjectSet の Sort で Nothing 要素を明示的に扱う
  - 詳細: `ObjectList` / `ObjectSet` は `Nothing` を通常要素として許可するが、`ObjectList.Sort` は `pIsComparable` / `pIsDuplicateCheckable` / `pIsStringable` の比較経路で `Me.Item(...)` を各インターフェイス変数へ `Set` するため、`Nothing` 要素を含むとオブジェクト変数未設定の実行時エラーになり得る。どの位置に並べるか、または Sort 対象外とするかの契約もない。
  - 影響: 集合・リストとしては追加、存在判定、削除できる `Nothing` が、ソート時だけ不安定に失敗する。`ObjectSet.Sort` は内部で `ObjectList.Sort` を使うため同じ影響を受ける。
  - 対応案: `Nothing` の順序を先頭/末尾のどちらかに固定するか、Sort では非対応として入口で明示エラーにする。`Nothing` のみ、`Nothing` と比較可能オブジェクト混在、`ObjectSet.Sort` のテストを追加する。

## 中優先度

- [ ] [bug] TextFileEntity のファイルモード・EOF・Close 契約を揃える
  - 詳細: `OpenFile` の `pAsAppend` 分岐で `If pGetReadLock And pGetReadLock Then` となっており、読み取りロック指定時に `Open` 文を実行しないまま `pIsOpen = True` へ進む可能性がある。`pGetWriteLock` との組み合わせ判定の誤記に見える。
  - 詳細: `IsEndOfFile` は `Output` / `Append` で開いた書き込み用ファイルでも `EOF(pFileDesc)` を呼ぶ一方、`TextFileEntityTestDouble.IsEndOfFile` は `Not pIsOpen Or pAsWrite` の場合に True を返しており、実装とテストダブルの契約がずれている。
  - 詳細: `TextFileEntity` は `OpenFile` で `FreeFile` により開いたファイルを、`CloseFile` が明示的に呼ばれた場合だけ閉じる。`Class_Terminate` での保険がなく、未オープン状態の close などを `Debug.Print` で扱う箇所もある。
  - 影響: 追記モード、ロック指定、書き込み用ファイルの EOF 判定、明示 Close 漏れが、実ファイルとテストダブルで異なる結果やファイルハンドル残りにつながる。
  - 対応案: `Input` / `Output` / `Append` と読み取り・書き込みロック、`CloseFile`、`Class_Terminate` の契約を表で仕様化する。実装とテストダブルに、各モード、Close 後、Append + lock、明示 Close 漏れのテストを追加する。

- [ ] [bug] TextFileEntity.Initialize の再初期化とオープン中呼び出しを拒否する
  - 詳細: `TextFileEntity.Initialize` は `pIsInitialized` や `pIsOpen` を確認せず、既存インスタンスの `pFilePath` をいつでも上書きできる。ファイルを開いた後に再初期化すると、`pFileDesc` は元のファイルを指したまま、`FilePath` やエラー文言だけが別パスになる。
  - 影響: 読み書き対象と診断情報がずれ、`CloseFile` や後続の `OpenFile` でどのファイルを扱っているか追跡しづらくなる。テストダブルやサービス層でインスタンス再利用した場合に、実ファイルハンドルの状態とオブジェクト状態が分離する。
  - 対応案: 初期化済みインスタンスの再初期化を明示エラーにし、必要なら `Reset` / 新規インスタンス生成を使う契約にする。未初期化、初期化済み、オープン中、クローズ後再初期化のテストを追加する。

- [ ] [bug] DiffStringArray の空配列入力と下限契約を修正する
  - 詳細: `DiffStringArray` は入力配列に対してすぐ `LBound` / `UBound` を呼ぶため、未初期化または空配列を渡すと差分結果ではなく実行時エラーになる。コメントでは `ChangeTypeArray` の下限を `OldArray` と揃えると読めるが、実装は常に `0 To ...` へ `ReDim` している。
  - 影響: 空ファイル、空行リスト、フィルタ結果 0 件の比較を共通 diff として扱えず、呼び出し側ごとに事前分岐が必要になる。下限を保つ前提の呼び出し側では添字ずれも起こり得る。
  - 対応案: 片側空、両側空、未初期化配列の扱いを仕様化し、出力配列の下限をコメントどおり維持するか、0 ベース固定としてコメントとテストを修正する。

- [ ] [bug] ExcelBookAndSheetAddress のクォート判定を Excel の参照構文へ合わせる
  - 詳細: `ExcelBookAndSheetAddress` はスペース、`'`、`!`、`[`、`]`、括弧だけをクォート対象にしているが、`入力-確認`、`2026.05`、先頭が数字のシート名など、Excel 数式上はクォートが必要な名前でも未クォートのアドレスを返し得る。
  - 影響: `RangeAddress`、`WorksheetRangeBounds.ToString`、`WorksheetService.XLookup` などが生成したアドレスや数式が、特定の実務シート名で評価できない可能性がある。
  - 対応案: Excel が未クォート参照として受け付ける名前をホワイトリストで判定し、それ以外は常にクォートする。シート名の先頭数字、ハイフン、ピリオド、日本語、クォートを含むケースのテストを追加する。

- [ ] [bug] FileSystemService のパス解決・バックアップ作成・コピー API を一括整理する
  - 詳細: `CommonModules/modules/FileSystemService.cls`: `CreateBackupFile` で `GetLeafFromPath(..., Extention:=...)` が使われているが、`Lib_Common.GetLeafFromPath` の引数名は `Extension`。
  - 詳細: `CopyDirectory` は `src_path` / `dst_path` を計算しているが `pFSO.CopyFolder` には正規化前の `SourcePath` / `DestinationPath` を渡している。`GetNewestFile` も `pFSO.GetFolder(DirectoryPath)` を直接呼び、他の一覧取得と相対パス基準が揃っていない。
  - 詳細: `IFileSystemService.cls`、`FileSystemService.cls`、`FileSystemServiceTestDouble.cls` の `CopyDirectory`、`MoveFile`、`CopyFile` は `DestinationPath` が `As String` なしの `Variant` になっている。
  - 詳細: 存在しないファイル・ディレクトリを `Debug.Print` で扱う箇所があり、呼び出し側が検知すべき失敗と無視可能な不存在の契約が曖昧になっている。
  - 影響: バックアップ作成経路がコンパイルまたは実行時に失敗し、相対パスや末尾セパレータ付きコピー先では存在確認と実操作の基準がずれる可能性がある。
  - 対応案: 名前付き引数を修正し、全ファイル/ディレクトリ操作で絶対パス化後の値を使う。インターフェイス、実装、テストダブルの型宣言を揃え、バックアップ名、相対パス、コピー先末尾セパレータ、`Force` 指定のテストを追加する。

- [ ] [bug] CounterSet の Dictionary 欠落キー参照を修正する
  - 詳細: `CounterSet.GetCounter` / `RemoveCounter` は `pDict(CounterObjectName)` を直接読み取っており、存在しない名前を指定すると `Scripting.Dictionary` が Empty 値を持つキーを新規作成してから `Set` 代入エラーになる。
  - 影響: 不存在取得が単なる失敗ではなく内部状態変更を伴うため、エラー捕捉後の後続処理やテストダブルの記録内容が実際とずれる。
  - 対応案: 取得前に `Exists` で存在確認し、不存在時は各クラスの明示的なエラーとして再送出する。`CounterSet` 欠落キー時に Count が増えないことを固定するテストを追加する。

- [ ] [bug] Lib_Common の値変換・配列・ソート系ユーティリティを整備する
  - 詳細: `SortArray` は再帰呼び出しで `Descending` を渡しておらず、比較関数名も `psortlessthan` / `pSortIsLessThan` で不一致になっている。昇順・降順の判定も期待と逆方向に見える。
  - 詳細: `IsInteger` / `IsLong` は `IsNumeric` 後に `CInt(Value)` / `CLng(Value)` を直接呼ぶため、`"32768"`、`"2147483648"`、`"1E+20"` のような範囲外の数値文字列で False を返す前に Overflow になる。
  - 詳細: `ConcatArray`、`ConvertArray2dTo1d`、`pSortSwap` はオブジェクト要素を扱う場合の `Set` 代入方針が揃っていない。`IsContainsIn` もオブジェクト配列で同一性・`IEquatable`・`IDuplicateCheckable` のどれを使うか未定義。
  - 対応案: 値変換は上下限と整数性を変換前に判定し、配列ユーティリティはオブジェクト要素の保持・比較方針を `ObjectList` / `ObjectSet` と揃える。数値境界、昇順/降順、空配列、オブジェクト配列のテストを `Test_Lib_Common.bas` に追加する。

- [ ] [bug] Lib_Common.GetMultiKey のキー衝突と特殊値変換を修正する
  - 詳細: `GetMultiKey` は引数を `Variant` で受ける一方、内部の `pGetMultiKeyEscape(ByVal DictionaryKey As String)` で文字列へ暗黙変換してから連結するため、`1` と `"1"`、`True` と `"True"` のように型が異なる値が同じキーになり得る。
  - 詳細: `Null` や `CVErr(...)` は文字列への暗黙変換で実行時エラーになり、複合キー生成そのものが失敗し得る。
  - 影響: 複数引数を Dictionary キーにする呼び出し側で、値の型差による取り違えや、Excel エラー値・Null を含むデータでの失敗が起こる。
  - 対応案: 値の型タグと特殊値表現を含むキー生成へ変更し、文字列・数値・Boolean・Empty・Null・CVErr を含む複合キーの衝突テストを追加する。

- [ ] [bug] 状態管理系小クラスの戻り値・境界値をテスト可能にする
  - 詳細: `Counter.Initialize` は `If CountStep = 0 Then` で既存プロパティを確認しており、引数 `CountStepNumber:=0` を検出できない。
  - 詳細: `DebugInformation.pBuildMessageByIndex` は先頭 `*` を設定した直後に上書きしており、`FinishTask(TaskName:=...)` は対象タスクを見つけても戻り値を `True` にしていない。
  - 詳細: `ProgressStatus.SetForLoop` は `pTotalValue = FinishIndex` を直接代入して `TotalValue` プロパティの検証を通らず、`Application.StatusBar` と `DoEvents` への依存により進捗計算単体のテストが書きにくい。
  - 対応案: `Counter`、`DebugInformation`、`ProgressStatus` の境界値、戻り値、表示文字列生成をテストで固定する。`ProgressStatus` は進捗率・文言生成と StatusBar 書き込みを分ける。

- [ ] [bug] SplitMessage の PageSize 境界値を検証する
  - 詳細: `SplitMessage` は `PageSize` を検証せず、`PageSize <= 0` の場合に `pTakeString` が 1 文字も消費しないため、長い行を処理する `Do While` が進まない。さらに `pTakeString` の `LengthByte` は `Integer` で、`SplitMessage` の `Long` 引数より狭い。
  - 影響: メッセージ表示系の補助 API に不正なページサイズが渡ると処理が戻らない、または大きなページサイズで型変換エラーになる可能性がある。
  - 対応案: `PageSize >= 1` を入口で検証し、`pTakeString` の `LengthByte` を `Long` に揃える。`PageSize` が 0、負数、32767 超、マルチバイト境界のテストを追加する。

- [ ] [bug] Lib_Common の配列境界・列挙ヘルパーの空配列と多次元配列契約を揃える
  - 詳細: `GetArrayBounds` は未初期化動的配列で 1 次元目の `LBound` / `UBound` が失敗すると、出力側の `LBoundArray` / `UBoundArray` も未初期化のまま返す。
  - 詳細: `GetArrayEnumerator` / `Enumerator` は `LBound` / `UBound` と `pArray(pIndex)` を 1 次元配列前提で使うため、`ReadRange` の戻り値のような 2 次元配列では列挙中に実行時エラーになる。
  - 影響: 空配列、未初期化配列、2 次元配列を共通配列ヘルパーへ渡したとき、入口で契約違反として落ちるのではなく、利用途中の別箇所で失敗する。呼び出し側ごとに防御分岐が増える。
  - 対応案: 1 次元配列専用 API と多次元配列対応 API を分け、サポート外は入口で明示エラーにする。空配列、未初期化配列、1 次元/2 次元配列のテストを追加する。

- [ ] [bug] Lib_IPv4.ParseIpAddressAndMask でマスク省略入力を明示的に扱う
  - 詳細: `G_IPV4_NW_RE` と `WellFormedAddress` はマスクなし IPv4 も扱えるように見えるが、`ParseIpAddressAndMask` は `Split(..., "/")` 後に `ip_arr(1)` を無条件参照する。`"192.168.0.1"` のようにマスク部がない入力では、形式エラーや `/32` 扱いではなく添字範囲外の実行時エラーになる。
  - 影響: IP アドレス単体をホスト `/32` として扱いたい呼び出しや、入力検証で正規表現を通した後の解析が、基盤 API 由来の分かりにくいエラーになる。
  - 対応案: `ParseIpAddressAndMask` 側でマスク省略を `/32` とするか明示的な形式エラーにする。`WellFormedAddress` との責務分担を決め、マスクなし、`/24`、`/255.255.255.0`、空マスクのテストを追加する。

- [ ] [bug] Lib_IPv4.NarrowNetwork の狭小化上限をコメントと実装で揃える
  - 詳細: `NarrowNetwork` のコメントは「マスク長が 32 に達している場合」にエラーと説明しているが、実装は `29 < MaskLength And MaskLength < 33` で `/30` と `/31` も拒否している。`/30` から `/31`、`/31` から `/32` を作るべきか、仕様として禁止するのかがコードから読み取れない。
  - 影響: IPv4 範囲を機械的に細分化する呼び出しで、まだ狭められるネットワークが基盤 API 側で止まる可能性がある。逆に `/31` や `/32` を扱わない方針なら、利用側がコメントを信じて境界条件を誤る。
  - 対応案: `/30`、`/31`、`/32` の扱いを仕様化し、実装・コメント・テストを揃える。ホストルート、ポイントツーポイント用途の `/31` を扱うかも合わせて決める。

- [ ] [bug] Lib_Common のビット演算・符号なし Long 演算の極端値を明示的に扱う
  - 詳細: `SubtractUnsignLong` は `ValueA < 0` かつ `0 <= ValueB` の分岐で `On Error Resume Next` を有効にしたまま `AddUnsignLong` を呼ぶため、補助関数側のエラーを握りつぶして後続計算へ進む可能性がある。
  - 詳細: `BitLeft` / `BitRight` は負のシフト数を逆方向シフトとして扱うが、`ShiftCount = -2147483648` のような `Long` 最小値では `-ShiftCount` 自体がオーバーフローする。32 ビット幅を超えるシフトを 0 埋めにするのか、明示エラーにするのかも固定されていない。
  - 影響: IPv4 計算など 32 ビット境界値を扱う基盤処理で、極端値だけ誤った結果、握りつぶされたエラー、または実行時エラーになる可能性がある。
  - 対応案: エラー制御を最小範囲に閉じ、シフト範囲と境界値の契約を入口で検証する。`0`、`31`、`32`、負数、`Long` 最小値、`&H80000000`、`&HFFFFFFFF` のテストを追加する。

- [ ] [spec] WorkbookService.OpenWorkbook の名称と実挙動を揃える
  - 詳細: `OpenWorkbook` は `Workbooks.Open` ではなく `Workbooks.Add(Template:=file_path)` でテンプレートから非表示ブックを作成している。コメントには元ファイルをロックしない旨があるが、API 名からは通常の open と誤解しやすい。
  - 影響: 呼び出し側が既存ブックを開いて操作する API と誤認し、保存先やロック、元ファイル更新の前提を取り違える可能性がある。
  - 対応案: 互換性を考慮しつつ `CreateWorkbookFromTemplate` / `OpenWorkbookCopy` など実態に近い名前を追加し、旧名はラッパーまたは非推奨コメントにする。

- [ ] [spec] TextFileEntity / TextFileService で文字コードと改行コードを指定できるようにする
  - 詳細: 現状は VBA の `Open ... For Input/Output/Append` と `Print #` に依存しており、UTF-8 BOM、Shift_JIS、CRLF/LF などを呼び出し側が明示できない。
  - 影響: 共通モジュールは UTF-8 BOM の `.md` / `.ps1` と Shift_JIS の `.bas` / `.cls` を扱うため、基盤テキスト API として文字コードの契約がないと使える範囲が狭い。
  - 対応案: 既存 API は互換維持し、必要に応じて encoding / newline を指定できるオプションまたは別サービスを追加する。代表的な文字コードと末尾改行のテストを追加する。

- [ ] [test] FileSystemServiceTestDouble.CreateDirectory の記録キーと戻り値キーを揃える
  - 詳細: `CreateDirectory` は呼び出し記録を `DirectoryPath, Force, Recursive` で残す一方、戻り値の取得は `DirectoryPath` だけをキーにしている。
  - 影響: `Force` や `Recursive` の違いで結果を切り替えるテストが書けず、実装が引数を正しく渡しているかをテストダブルで表現しにくい。
  - 対応案: 記録、戻り値、既定値取得のキー生成を同じ引数セットに揃える。既存テストへの影響を確認し、必要なら互換用の設定ヘルパーを追加する。

- [ ] [ref] WorksheetRangeBounds の値オブジェクト化と境界仕様をまとめる
  - 詳細: `WorksheetRangeBounds.Initialize` は `Book = ""` のとき `New WorkbookService` で `ThisWorkbook.Name` を取得しており、範囲値オブジェクトが Excel サービス生成を内包している。
  - 詳細: `TransformAbsolute`、`Transform` は `New_RangeBounds` ファクトリを呼ぶため、`Lib_Common` と `WorksheetRangeBounds` の相互依存ができている。`UserInputSheet.GetItemRange` では `TransformAbsolute` が「使用範囲の先頭列を取る」用途にも使われている。
  - 詳細: `Initialize` は `FinishRow` / `FinishColumn` を `G_ROW_MAX` / `G_COL_MAX` に丸める一方、開始 `Row` / `Column` が上限を超えた場合は丸めもエラー化もしない。`Count` の桁あふれは中優先度の `[req] WorksheetRangeBounds.Count のセル数桁あふれを防ぐ` に切り出す。
  - 対応案: 既定ブック名の補完は `New_RangeBounds` / `Constructor.bas` 側へ寄せ、`WorksheetRangeBounds` は渡された値の正規化と検証だけを担当する。開始行・開始列の上限超過、巨大範囲の `Count`、`Transform` 系の用途をテストで固定する。

- [ ] [ref] Err.Raise の Source と説明文を診断しやすい表記へ揃える
  - 詳細: `ConvertStringToBoolean` のエラー `Source` が `Sub CreateBackupFile` になっている。`pA1columnAddressCore` は列インデックスの検証で「行インデックス」、`pA1RowAddressCore` は行インデックスの検証で「列インデックス」と表示する。
  - 詳細: `Err.Raise vbObjectError + 1, "Class ...", ...` とキーワード指定ありの書き方が混在しており、標準モジュールでは `Function` / `Sub` の粒度も揺れている。
  - 影響: 利用側で例外を再送出・表示したときに、失敗箇所や対象パラメータを誤認しやすい。
  - 対応案: `Source:="Function Xxx"` / `Source:="Class Xxx"` の規則とメッセージ内の引数名を揃え、`Test_ErrSource.bas` の対象を標準モジュールの代表関数にも広げる。

- [ ] [ref] WorksheetService のシート単位操作と範囲単位操作を分ける
  - 詳細: `SetAllDataVisible` と `SetSheetOutlineLevel` は `RangeBounds` を受け取るが、実際には対象シート全体のフィルタ、非表示行列、アウトラインを変更する。`InsertRows` / `DeleteRows` も範囲内セルではなく行全体を操作するため、メソッド名だけでは変更範囲を読み取りにくい。
  - 影響: 基盤サービスとして呼び出す側が、範囲に限定した操作だと誤解してシート全体の表示状態や行構造を変えてしまう可能性がある。
  - 対応案: シート単位 API は `WorksheetDisplayService` などへ分離するか、`SetSheetAllDataVisible` / `InsertEntireRows` のように作用範囲が分かる名前へ段階移行する。既存名は互換ラッパーとして残す。

- [ ] [spec] WorksheetService.XLookup の Excel バージョン依存と代替方針を明示する
  - 詳細: `WorksheetService.XLookup` は Excel の `XLOOKUP` 関数を式文字列として評価するため、関数が利用できない Excel では基盤 API 自体が使えない。戻り値が Excel エラー値になった場合も、未検出、評価失敗、関数未対応の区別が呼び出し側から分かりにくい。
  - 影響: 共通モジュール利用側が Excel バージョンや関数対応状況を意識せずに呼ぶと、環境差で検索処理だけ失敗する可能性がある。
  - 対応案: `XLookup` をバージョン依存 API としてコメントで明示し、必要なら `Match` / `Index` ベースの互換検索 API を追加する。関数未対応、未検出、評価エラーを区別するテストまたは戻り値契約を用意する。

- [ ] [spec] WorksheetRangeBounds.Count のセル数桁あふれを防ぐ
  - 詳細: `WorksheetRangeBounds.Count` は `RowCount * ColumnCount` を `Long` で返すため、シート全体や大きな範囲では `1,048,576 * 16,384` のように `Long` 上限を超えて桁あふれする。
  - 詳細: `WorksheetRangeBoundsEnumerator.Initialize(EnumType:="Cells")` も `pRangeBounds.Count` を `Long` の `pLength` へ代入するため、全シート級のセル列挙を始める前に失敗し得る。
  - 影響: 範囲オブジェクトの基本プロパティ参照やセル列挙が、大きな範囲だけ実行時エラーになり、基盤 API として安全に扱えない。
  - 対応案: 対応可能な最大セル数を仕様化し、桁あふれ前に明示エラーにするか、戻り値型・列挙契約を見直す。全シート、全列、`Long` 上限付近の範囲テストを追加する。

- [ ] [spec] WorksheetRangeBounds.GetEnumerator で降順列挙を公開 API から指定できるようにする
  - 詳細: `WorksheetRangeBoundsEnumerator.Initialize` は `Descending` を受け取り、列挙子本体も降順移動に対応しているが、`WorksheetRangeBounds.GetEnumerator` の公開引数には `Descending` がない。行・列・セルの列挙方向指定と並び順指定の入口が揃っていない。
  - 影響: 呼び出し側が範囲を下から上、右から左へ走査したい場合、公開 API からは指定できず、列挙子クラスを直接生成する必要がある。基盤型としての使い勝手が落ちる。
  - 対応案: `GetEnumerator(Optional EnumerateType..., Optional ColumnDirection..., Optional Descending As Boolean = False)` に拡張するか、`GetDescendingEnumerator` のような別 API を追加する。既存呼び出し互換を保つ引数順に注意する。

## 低優先度

- [ ] [spec] WorksheetService.Sort のキー指定を型付きで表現する
  - 詳細: `Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)` は、キー列番号と並び順を交互に並べる呼び出し規約になっている。呼び出し側では `Sort bounds, 2, xlAscending, 1, xlDescending` のような値の意味が読み取りにくく、テストダブルも生の配列として保持している。
  - 影響: キー列番号と並び順の入れ替え、片方だけの指定漏れ、既定値の解釈違いが起こりやすい。将来、キーごとのオプションを増やす場合も ParamArray の位置規約がさらに複雑になる。
  - 対応案: `WorksheetSortKey` や `WorksheetSortSpec` のような型、または `New_SortKey(ColumnIndex, Order)` 形式の薄いファクトリを追加し、既存 ParamArray API は互換ラッパーとして扱う。

- [ ] [spec] WorksheetService.EvaluateFormula の ContextRangeBounds 契約を明確にする
  - 詳細: `EvaluateFormula(ByVal ContextRangeBounds, ByVal Formula)` は `ContextRangeBounds` から対象シートだけを取り出し、`target_sheet.Evaluate(Formula)` を呼ぶ。引数名からはセル位置を含む評価コンテキストに見えるが、相対参照や R1C1 参照をどのセル基準で評価するかは契約化されていない。
  - 影響: 呼び出し側が `ContextRangeBounds` の左上セルを基準に評価されると期待すると、式の書き方によっては結果がアクティブ状態や Excel の評価仕様に依存して読み取りにくくなる。
  - 対応案: シートコンテキストだけを扱うなら `EvaluateOnWorksheet` などへ改名またはコメント明記する。セル基準評価を扱うなら、対象セルの `Range.Evaluate` 相当または一時セル評価などの方針を決め、A1 / R1C1 / 名前参照のテストを追加する。

- [ ] [ref] Tmp_ManualTest の残存コードを整理する
  - 詳細: `Tmp_ManualTest.bas` は `Option Explicit` がなく、個人環境の絶対パスや `Debug.Print` を含む手動確認コードが残っている。対応済み欄の「手動検証用モジュールを整理する」では `Module1.bas` の削除のみが最終対応になっており、この残件が open 項目として追跡されていない。
  - 影響: 共通モジュールの import/export や静的確認で、実行対象ではない個人用コードが混ざり、規約違反や依存関係調査のノイズになる。
  - 対応案: 残す必要がある手動検証は `tool` 側またはドキュメントへ退避し、VBA モジュールとして残す場合は `Option Explicit`、環境依存パスの排除、実行入口の明確化を行う。

- [ ] [ux] ObjectList / ObjectSet を標準の For Each で列挙できるようにする
  - 詳細: `ObjectList` / `ObjectSet` は独自の `IEnumerator` を返すが、VBA 標準の `_NewEnum` には対応していない。
  - 影響: 基盤コレクションとしては利用者が期待する `For Each` が使えず、簡単な読み取り列挙でも独自列挙子の書き方を覚える必要がある。
  - 対応案: 既存 `IEnumerator` を維持しつつ、読み取り用の標準列挙を追加できるか検討する。追加する場合は列挙中更新との契約を明記する。

- [ ] [ref] ConvertRangeToStringList のグローバルサービス依存と読み取り契約を整理する
  - 詳細: `ConvertRangeToStringList` は `WsSrv.ReadCell` を直接呼ぶが、関数内で `InitializeCommonService` などの初期化は行わない。コメントは Range の Text を文字列化すると読めるが、`ReadCell` の既定値では Value 側を読む。
  - 影響: 呼び出し順によって未初期化の `WsSrv` に依存し、表示文字列を取りたい利用者にも値文字列を返すため、基盤関数として挙動を読み取りにくい。
  - 対応案: `IWorksheetService` を引数で受ける、または初期化済みサービス前提をコメントで明示する。Text/Value のどちらを返すかを名前または引数で分ける。

- [ ] [spec] Lib_Common のパス・Excel アドレス・ファイル形式ヘルパー境界仕様を固定する
  - 詳細: `GetParentPath("file.txt")` のように区切り文字を含まないパスでは `InStrRev` が 0 となり、`Left(Path, -1)` で失敗する。ドライブ直下、UNC ルート、末尾セパレータ、拡張子だけのファイル名も仕様が明確でない。
  - 詳細: `GetParentPath` / `GetLeafFromPath` は位置計算に `Integer` を使っている箇所があり、長いパスではオーバーフローする可能性がある。
  - 詳細: `RangeAddress` は A1 形式の絶対参照指定で不要に見える `ReferenceRow` / `ReferenceColumn` を要求し、列範囲の終了列では `IsAbsoluteFinishColumn` ではなく `IsAbsoluteStartColumn` を渡している。`ExcelA1ColumnAddress` は 0 以下や Excel 最大列超過の仕様が実装とコメントでずれている。
  - 対応案: `JoinPath`、`GetParentPath`、`GetLeafFromPath`、`IsAbsolutePath`、`RangeAddress`、`ExcelA1ColumnAddress`、`GetExcelFileFormat` の境界値を `Test_Lib_Common.bas` で固定し、その後に `Lib_Path` / `Lib_ExcelAddress` へ分離する。

- [ ] [ref] 範囲走査と ReadRange 結果処理を統一する
  - 詳細: `フォルダー生成/modules/Mod_ReplacePlaceholder.bas`、`申請書ツール/modules/Mod_MultiRowFormat.bas`、`重複ネットワーク生成/modules/GUIHandler.bas` に、`WorksheetRangeBounds` の `Row` / `FinishRow`、`Column` / `FinishColumn` を手動走査する処理がある。
  - 詳細: `CES勤務表/modules/CustomHolidayObject.cls` と `差分チェック/modules/Mod_DiffColumns.bas` は `ReadRange` の戻り配列を直接 `LBound` / `UBound` で走査している。
  - 詳細: `コンフィグ生成/modules/Mod_InsertRow.bas` の `RowRangeInformation.StartRow To FinishRow` は直接の `WorksheetRangeBoundsEnumerator` 対象ではないが、行範囲表現として似た整理余地がある。
  - 対応案: セルや行を意味として処理している箇所は `WorksheetRangeBoundsEnumerator` へ寄せる。一括 `ReadRange` の性能が主目的の箇所は、1 列範囲を配列へ変換するヘルパーなどを検討する。

- [ ] [ref] 共通モジュールコピーの同期ずれを検出・解消する
  - 詳細: `common_modules_repo` と各プロジェクトの同名ファイルをハッシュ比較すると、`申請書ツール/modules/IWorkbookService.cls` と `WorkbookServiceTestDouble.cls` は末尾改行なし、`申請書ツール/modules/Lib_Common.bas` はコメント内空白差分、`申請書ツール/ArrayObject.cls` は modules 外の古いコピーと思われる差分がある。
  - 影響: 共通モジュールのリファクタリング時に、同期対象外のコピーや末尾改行差分がノイズになり、どれが正本か判断しにくくなる。
  - 対応案: `import_common_main.ps1` 実行後の検証として、`CommonModules/modules`、`common_modules_repo`、各プロジェクト `modules` の同名ファイルをハッシュ比較するスクリプトまたはテストを追加する。modules 外の stray copy は削除可否を確認して整理する。

- [ ] [ref] sync.json 対象モジュールの正本とテスト配置を整理する
  - 詳細: `申請書ツール/sync.json` では `申請書ツール` と `コンフィグ生成` の PAN 系 17 モジュールを同期しており、ハッシュ上も同一だが、CommonModules のような単独の正本・テストブック・配布元ディレクトリはない。
  - 詳細: `sync_modules_main.ps1` は対象プロジェクト群のうち更新日時が最も新しいファイルを採用するため、意図せず片側の作業中コピーを正として同期するリスクがある。
  - 詳細: PAN 系の値オブジェクト、読み書き、コンストラクターは複数ツールで共有される実質的な共通ドメイン層だが、共通モジュールの TODO やテスト実行手順とは別管理になっている。
  - 対応案: 同期モジュールにも `shared_modules_repo` 相当の正本ディレクトリまたは専用テストブックを用意する。少なくとも同期前後のハッシュ確認と、どちらのプロジェクトでテストするかを TODO/手順に明記する。

- [ ] [ref] Lib_Common を責務別モジュールへ分割する
  - 詳細: `Lib_Common.bas` はサービス初期化、ファクトリ、エラー処理、GUI ボタン、クリップボード、パス操作、配列操作、文字列処理、diff、ビット演算、Excel アドレス生成が同居している。
  - 詳細: `InitializeCommonService`、`New_RangeBounds`、`AddButton`、`GetLeafFromPath`、`ConvertArray...`、`DiffStringArray`、`RangeAddress` などに `Public` / `Private` の明示がない宣言が多く、公開 API と内部ヘルパーの境界が読み取りにくい。
  - 詳細: `FileSystemService.cls` と `WorkbookService.cls` の `pGetAbsolutePathCore`、`FileSystemService.cls` / `WorkbookService.cls` / `WorksheetService.cls` の `pAddItemToStringArray` など、共通化できる小ヘルパーの重複も残っている。
  - 対応案: `Constructor.bas`、`Lib_Array.bas`、`Lib_String.bas`、`Lib_Path.bas`、`Lib_ExcelAddress.bas`、`Lib_Error.bas`、`Lib_Clipboard.bas` などへ段階分割する。先に現利用箇所とワークシート関数利用の有無を棚卸しし、互換ラッパーを残す範囲を決める。

- [ ] [ref] WorksheetService / IWorksheetService を責務別インターフェイスへ分割する
  - 詳細: `WorksheetService.cls` と `IWorksheetService.cls` は Find、Sort、入出力、式評価、コピー、UsedRange、行削除、書式設定まで 1 インターフェイスに含む。`WorksheetServiceTestDouble.cls` も公開 Dictionary が 20 個以上になっている。
  - 詳細: 読み取りだけ必要な処理でも、書式・コピー・行操作まで含む巨大な `IWorksheetService` へ依存するため、テストダブルの更新範囲が広がりやすい。
  - 対応案: 互換性を保ちながら `IWorksheetReader`、`IWorksheetWriter`、`IWorksheetFormatter`、`IWorksheetRangeFinder` などへ段階分割する。既存 `WsSrv As IWorksheetService` は移行期間の facade とし、テストダブルも小さな部品を合成する形へ寄せる。

- [ ] [test] CommonModules のテスト基盤と未テスト領域を整理する
  - 詳細: `Lib_UnitTest.bas` は `ThisWorkbook.VBProject` を直接走査し、`Test_ErrSource.bas` は `Application.VBE.ActiveVBProject` を走査している。実行時にアクティブプロジェクトが変わると、意図しないプロジェクトを検査する可能性がある。
  - 詳細: テスト検出・実行・結果シート書き込み・再実行ボタン追加が `Lib_UnitTest.bas` にまとまっており、テストランナー単体の確認がしにくい。
  - 詳細: `Test_*.bas` が直接対応していない共通モジュールに、`ArrayObject.cls`、`Counter.cls`、`CounterSet.cls`、`DebugInformation.cls`、`Enumerator.cls`、`Lib_IPv4.bas`、`ObjectSet.cls`、`UnitTestAssert.cls`、`UnitTestUtils.cls`、`UserInputSheet.cls` がある。
  - 対応案: VBProject 走査、テスト候補抽出、結果出力を小さな関数に分ける。未テスト領域は、実バグが見つかっている `ObjectSet`、`Counter`、`UnitTestAssert`、`DebugInformation`、実務ロジック境界を担う `Lib_IPv4` から追加する。

- [ ] [ux] UnitTestAssert.EqualsArray の差分位置を多次元座標で表示する
  - 詳細: `EqualsArray` は多次元配列の境界は比較するが、要素差分の位置は `For Each` の線形インデックス `@0`、`@1` ... として出力する。`ReadRange` のような 2 次元配列では、どの行・列の不一致かを結果メッセージから直接読めない。
  - 影響: 2 次元の表データ比較で NG になったとき、線形位置から行列位置へ手作業で戻す必要があり、テスト失敗調査に時間がかかる。
  - 対応案: 配列の次元と LBound / UBound を使って、差分位置を `@(row=..., col=...)` のように表示する。1 次元、2 次元、非 0 ベース配列の失敗メッセージをテストで固定する。

- [ ] [ux] UnitTestAssert のエラーアサーション API を呼び出しやすくする
  - 詳細: `ErrorRaised` / `ErrorNotRaised` は他のアサーションと異なり `Function` として Boolean を返しつつ、内部の失敗状態も更新する。呼び出し側は `Err.Number` / `Err.Source` / `Err.Description` を個別に退避して渡す必要があり、`On Error Resume Next` と `Err.Clear` の定型コードがテストごとに散る。
  - 影響: エラー発生を確認するテストの形が統一されず、戻り値をさらに `Assert.IsTrue` するのか、呼び捨てでよいのかも読み取りにくい。エラー捕捉コードの書き漏れで、直前の `Err` が混ざるリスクもある。
  - 対応案: 既存 API の互換性を保ちながら、エラー捕捉とアサーションをまとめるヘルパー、または `Sub` 形式の明示的なアサーション名を追加する。既存コメントの `ActualCodeCode` typo も合わせて直す。

- [ ] [test] ObjectList の多段拡張境界をテストする
  - 詳細: `ObjectList` は 16 要素固定の `ArrayObject` を多段化して可変長を実現しているが、既存テストは主に 20 件程度までで、`256`、`4096` など複数段拡張の境界を直接確認していない。
  - 詳細: `Remove` は要素を前詰めし、内部ツリーは縮小しないため、多段拡張後の削除・更新・コピー・列挙が境界をまたいで正しく動くかも固定したい。
  - 影響: 深い階層でのインデックス分解や root 拡張に不具合があっても、現状のテストでは検出しにくい。
  - 対応案: `15` / `16` / `17`、`255` / `256` / `257`、`4095` / `4096` / `4097` 付近の `Add` / `Item` / `Update` / `Remove` / `CopyList` / 列挙をテストで固定する。

- [ ] [spec] ObjectSet のインデックス境界を明示エラーにする
  - 詳細: `ObjectSet.Item` / `Update` / `Remove` は範囲外インデックスの明示チェックがなく、`pDict.Keys(Index)` などの Dictionary 側のエラーに依存している。
  - 影響: 空集合、負数、`Count` 以上の指定時に、エラーの発生元と説明が呼び出し側から追跡しづらい。
  - 対応案: 有効範囲を `0 <= Index < Count` として `ObjectSet` 自身のエラーを送出し、`Item` / `Update` / `Remove` の境界テストを追加する。

- [ ] [spec] WorksheetRangeBounds.Item のインデックス基準を他メソッドと揃える
  - 詳細: `WorksheetRangeBounds.Item` は 0 オリジンでセルを返す一方、`GetRow` / `GetColumn` / `GetCell` は 1 オリジンの `RowIndex` / `ColumnIndex` を受ける。`Item` のコメントは「番号」とだけ書かれており、基準の違いが公開 API から読み取りにくい。
  - 影響: 範囲を直接走査する利用者が `Item(1)` を先頭セルと誤解すると、2 番目のセルから処理してしまう。`WorksheetRangeBoundsEnumerator` 経由では 0 オリジンが内部実装として隠れているため、直接利用時だけ罠になる。
  - 対応案: 既存 `Item` の互換性を保つなら、コメントと引数名で 0 オリジンを明示する。1 オリジンで公開する API が必要なら、`GetCellByIndex` など別名で追加し、`Item` は列挙子向けの低レベル API と位置付ける。

- [ ] [ref] WorksheetRangeBoundsEnumerator.Initialize の TargetCollection 引数名を範囲に合わせる
  - 詳細: `Initialize(ByVal TargetCollection As WorksheetRangeBounds, ...)` は実際には `WorksheetRangeBounds` を受け取るが、引数名とコメントは汎用列挙子と同じ「コレクション」になっている。範囲値オブジェクトを渡す専用列挙子であることが公開 API から伝わりにくい。
  - 影響: 呼び出し側やテストで、配列や `ObjectList` 用の汎用 `Enumerator` と同じ契約だと誤読しやすい。`WorksheetRangeBounds` 専用列挙子として、どの型を列挙対象にするのかも読み取りにくい。
  - 対応案: named argument 互換性を考慮しつつ、`TargetRangeBounds` など範囲であることが分かる名前やファクトリ API へ移行する。少なくともコメントは「列挙対象の範囲」に修正し、汎用 `Enumerator` との命名差を整理する。

- [ ] [ref] WorkbookServiceTestDouble の CloseWorkbook 記録名とコメントを実 API 名へ揃える
  - 詳細: `WorkbookServiceTestDouble` の公開 Dictionary が `ColoseWorkbook_Results` という綴りになっており、`CloseWorkbook` と一致していない。周辺コメントにも `SaveWorkbook` や `RemoveWorksheet` の説明が残っている。
  - 影響: テストで記録結果を確認するときに自然な `CloseWorkbook_Results` という名前を使えず、テストダブルの公開 API に typo を広げることになる。コメントからもどのメソッドのスパイ結果か読み取りづらい。
  - 対応案: 互換性が必要なら旧名を段階廃止扱いにしつつ、正しい `CloseWorkbook_Results` へ移行する。コメントと既存テストの参照名も合わせて更新する。

- [ ] [spec] ObjectList / ObjectSet の空集合化後の型固定契約を決める
  - 詳細: `ObjectList.Remove` や `ObjectSet.RemoveItem` / `Remove` で要素数が 0 になっても、`pTypeName` や `pIsObject` などの型状態は残る。一方で `RemoveAll` は型状態をリセットするため、空になった経路によって次に追加できる型が変わる。
  - 影響: 利用者から見ると同じ空のコレクションでも、個別削除後は過去の型に縛られ、`RemoveAll` 後は新しい型を受け入れる。値コレクション基盤として、空集合の再利用可否が操作履歴に依存して分かりにくい。
  - 対応案: 空になっても型固定するのか、最後の要素削除時に型状態もリセットするのかを仕様化する。`Remove` / `RemoveItem` / `RemoveAll` 後に別型を追加するテストを `ObjectList` / `ObjectSet` 両方へ追加する。

- [ ] [ref] 共通サービス初期化とテストダブル記録・キー生成を統一する
  - 詳細: `Lib_Common.bas` の `WbSrv` / `WsSrv`、`Lib_FileSystem.bas` の `FsSrv`、`Lib_TextFile.bas` の `TfSrv` が別々にグローバル管理され、`UserInputSheet.cls` などでは直接 `New WorkbookService` / `New FileSystemService` している箇所もある。
  - 詳細: `WorksheetServiceTestDouble.cls`、`WorkbookServiceTestDouble.cls`、`FileSystemServiceTestDouble.cls`、`TextFileEntityTestDouble.cls` は、`Public Xxx_Values As Dictionary` / `Public Xxx_Results As Dictionary` と `UnitTestUtils.GetValue` / `SetValue` の組み合わせを各メソッドで繰り返している。
  - 詳細: `WorkbookServiceTestDouble.cls` の `ColoseWorkbook_Results` など、公開フィールド名の誤記がテスト側 API として固定化されている。
  - 詳細: `UnitTestUtils.pGetKeyCore` はテストダブルの各メソッドから個別に呼ばれており、キー生成規則が `ObjectList` / `ObjectSet` などの値比較方針と独立している。特殊値での実行時バグは高優先度の `UnitTestUtils の引数キー生成で Null・CVErr を安全に扱う` に切り出す。
  - 詳細: オブジェクトは `IEquatable` なら `GetIdentityString`、それ以外は `ObjPtr` をキーにしており、`ObjectList` / `ObjectSet` の重複判定方針と似ているが独立実装になっている。
  - 対応案: サービス取得関数を 1 箇所にまとめ、将来的には `ServiceProvider` 的な小クラスで既定実装とテストダブル差し替えを明示する。テストダブルの記録処理とキー生成は `TestDoubleCallStore` / `VariantKeyBuilder` のような共通クラスへ寄せ、サポート対象外の引数型では明示エラーにする。

- [ ] [ref] Lib_IPv4 の仕様境界とテストを CommonModules に移す
  - 詳細: `Lib_IPv4.bas` は `申請書ツール`、`重複ネットワーク生成` などの実処理で使われるが、CommonModules 側に `Test_Lib_IPv4.bas` がない。
  - 詳細: `G_IPV4_NW_RE` はマスク省略を許す形だが、`ParseIpAddressAndMask` は `ip_arr(1)` が存在する前提で動く。`/0`、`/32`、負値になる IPv4 Long、ドットマスクとマスク長の相互変換など、仕様境界をテストで固定したい。
  - 詳細: `NarrowNetwork` はコメントでは「マスク長が 32 に達している場合」をエラーとしているが、実装は `29 < MaskLength And MaskLength < 33` で `/30` と `/31` も拒否する。
  - 対応案: まず現仕様をテスト化し、ネットワーク表記の正規化、パース、マスク変換、ネットワーク拡張・分割を `AddressObject` などツール固有処理から切り離して検証する。

- [ ] [ref] 未使用または低利用の公開 API を棚卸しする
  - 詳細: 共通モジュール名と同名のコピーを除いて静的参照を数えると、`Lib_Common.DIFFSTR`、`AddButton`、`DeleteButton`、`SetClipboard`、`GetClipboard`、`PasteFormulas`、`SortArray`、`LongToBin`、`BitLeft`、`BitRight`、`WorksheetService.WriteArrayFormula`、`WorkbookService.AddWorksheet` などに通常コードからの参照がほぼ見当たらない。
  - 詳細: `PasteFormulas` は `Selection` / `ActiveSheet` と広い `On Error Resume Next` に依存する UI 操作用 API なので、残す場合も `Lib_Clipboard` 側へ分離し、対象範囲を明示できる形にしたい。
  - 注意: ワークシート関数、ボタン割り当て、外部ブックからの呼び出しは静的検索だけでは検出できない。
  - 対応案: 削除ではなく、まず公開理由を `@details` または TODO に記録する。利用実績がないものは `Deprecated` コメントや互換ラッパーを挟んでから整理する。

- [ ] [ref] Lib_Common 先頭の未使用コメントアウトコードを整理する
  - 詳細: `Lib_Common.bas` 先頭に、未実装コメントを含む `TEXTSPLIT` / `TEXTJOIN` 互換関数のコメントアウト済みコードが大きく残っている。残す場合は正式なワークシート関数としてテスト化し、不要なら削除して履歴参照に寄せたい。
  - 対応案: 残す API と削除する履歴コメントを分ける。`TEXTSPLIT` / `TEXTJOIN` 互換関数を正式化する場合は仕様、公開コメント、単体テストを追加し、不要ならコメントアウト塊を削除する。

- [ ] [ref] Option Base と内部用手続きのスコープを規約へ揃える
  - 詳細: 配列操作を含む一部モジュールに `Option Base 0` がない。`UnitTestAssert.pEqualsCore` や `ObjectList.pInitializeItemType` / `ObjectSet.pInitializeItemType` のように、`p` 接頭辞の内部用手続きが `Public` または暗黙 Public になっている箇所がある。
  - 対応案: 配列下限に依存するモジュールへ `Option Base 0` を追加し、内部用手続きは `Private` を明示する。外部から呼ばれていた可能性があるメンバーは参照を棚卸しして、必要なら互換ラッパーを残す。

- [ ] [ref] 公開 API とコメントの綴り揺れを段階整理する
  - 詳細: `UpperBonnd`、`ColoseWorkbook_Results`、`pCalcurateArrayCount`、`WorkgbookService` / `WorkgsheetService` など、公開 API やコメントに綴りの揺れがある。`ArrayObject.Update` は値を更新するだけだが `Public Function` かつ戻り値型なしで宣言されている。
  - 詳細: `CompareAsUnsignLong`、`IsLessThanUnsignLong`、`AddUnsignLong`、`SubtractUnsignLong` は、一般的な `Unsigned` ではなく `Unsign` になっており、公開 API として検索しづらい。
  - 対応案: 公開名の誤記は互換性を壊すため、旧名を残すプロパティ・ラッパーを置いて段階移行する。コメントだけの誤記は挙動変更なしで修正し、`ArrayObject.Update` は呼び出し元確認後に `Sub` 化または戻り値型明示を検討する。`Unsign` 系 API は `Unsigned` 表記のラッパーを追加して利用側を移行し、旧名は互換維持または非推奨コメントを付ける。

- [ ] [ref] Doxygen コメント形式と公開メンバー記述を統一する
  - 詳細: `FileSystemService.cls` の `CreateBackupFile`、`TextFileEntity.cls` の公開 API コメントなど、Doxygen 対象の `'* ` 形式になっていない箇所がある。`@return` タグ、`Optional` の書き方、Public メンバーへの `@details` 有無にも揺れがある。
  - 対応案: 公開 API コメントを `'* `、`@return`、`@details` ありの規約へ揃える。コメントのみの変更として import/export 後の差分を確認し、必要ならドキュメント生成結果も確認する。

- [ ] [ref] ApplicationScreenUpdateManager のスコープ管理と復元契約を明確にする
  - 詳細: `ApplicationScreenUpdateManager.cls` は生成時に Excel 設定をバックアップし、`Restore` 後に `pIsBackedUp = False` へ戻すため、同一インスタンスを再利用して `DisableUpdates` する用途ではエラーになる。
  - 詳細: 複数インスタンスをネストした場合、LIFO で破棄される前提なら動くが、途中で片方を明示 `Restore` した場合や、他処理が `Application` 設定を変更した場合の契約が明記されていない。
  - 詳細: `ProgressStatus` は `Application.StatusBar` を直接 `False` に戻しており、画面更新・イベント・計算・警告とは別系統で Excel 状態を復元している。直前のステータスバー文字列を保持する契約もない。
  - 対応案: `ApplicationStateScope` のように復元対象を明示できる小クラスへ整理し、再利用可否、二重 `Restore`、ネスト、StatusBar の復元方針をテストで固定する。

- [ ] [ref] params シート読み取りを共通の設定リーダーへ切り出す
  - 詳細: `申請書ツール/modules/ParamObject.cls`、`フォルダー生成/modules/ParamObject.cls`、`セキュリティポリシー検索/modules/ParamObject.cls` は、`params` シートの名称列を検索して値列を読む `GetParam` ロジックをそれぞれ持っている。
  - 詳細: 実装ごとに `InitializeCommonService` の有無、検索範囲、エラー文言、定数名が揺れている。各プロパティ呼び出しのたびに UsedRange と Find を実行するため、同じパラメータを繰り返し読む処理では無駄が出る。
  - 詳細: 戻り値はほぼ文字列で、数値列・パス・CSV リストなどの型変換が各利用側へ散っている。
  - 対応案: CommonModules に `ParameterSheetReader` / `ParameterStore` 相当を追加し、シート名、名称列、値列、開始行を指定して辞書へキャッシュする。ツール固有の `ParamObject` はキー名付きの薄いプロパティラッパーだけにする。

- [ ] [ref] coding_rule_violations.md の公開名残件を TODO 側へ統合する
  - 詳細: `coding_rule_violations.md` には、`ApplicationScreenUpdateManager`、`WorksheetRangeBoundsEnumerator`、各 `Test_*TestDouble` など推奨 26 文字を超えるモジュール名と、`FsSrv` / `TfSrv` の公開グローバル変数名が残件として別管理されている。
  - 影響: TODO と規約違反一覧が分かれているため、命名整理の優先順位や完了条件を追いにくい。
  - 対応案: 互換性影響がある公開名とテスト内部名を分け、リネームするもの、互換ラッパーを残すもの、既存名として許容するものを TODO 側で管理する。

- [ ] [ref] ObjectList / ObjectSet / ArrayObject の名前と公開範囲を値コレクションとして整理する
  - 詳細: `ObjectList` / `ObjectSet` は名前に `Object` を含むが、実際にはプリミティブ値、`Nothing`、一部の特殊値も扱う設計になっている。`ArrayObject` は `ObjectList` の内部実装に近いが、公開クラスとして見えており、基盤利用者が直接使うべき型か判断しにくい。
  - 影響: コレクション基盤として、値型も入れてよいのか、オブジェクト専用なのか、内部用クラスを直接使ってよいのかが名前から伝わりにくい。
  - 対応案: 既存名は互換維持しつつ、`VariantList` / `VariantSet` / `ArrayBuffer` など実態に近い名前の追加や、`ArrayObject` の内部用コメント・非推奨コメントを検討する。

- [ ] [ref] 判定・取得系 API の英語名を一般的な形へ揃える
  - 詳細: `GetAllWorkbook` / `GetAllWorksheet` は戻り値が配列なのに単数形で、`ExistsWorkbook` / `ExistsWorksheet` は英語としては `WorkbookExists` / `WorksheetExists` などの方が自然に読める。`IsContainsIn`、`ObjectSet.GetContains`、`ConvertArrayStringToVariant` / `ConvertArrayVariantToString` も、動詞・目的語・戻り値の関係が読み取りにくい。
  - 影響: 共通基盤 API として、初見の利用者が検索しづらく、戻り値や副作用を名前から推測しにくい。類似 API を追加すると命名の揺れが広がりやすい。
  - 対応案: 既存名は互換ラッパーとして残し、`GetAllWorkbooks`、`WorkbookExists`、`ArrayContains`、`GetContainedItem`、`ConvertStringArrayToVariantArray` などの移行先名を決める。利用箇所を静的検索して段階的に置き換える。

- [ ] [ref] 変換ヘルパーの引数型・戻り値型を名前と揃える
  - 詳細: `ConvertBooleanToString` は Boolean 変換なのに `BooleanValue As String`、`ReplaceSpecialCharacterOnFileSystemPath` は常に文字列を返すのに戻り値が `Variant`、`RangeAddress` も文字列生成なのに戻り値が `Variant` になっている。
  - 影響: 呼び出し側が意図しない型変換に依存しやすく、テストや補完から契約を読み取りにくい。
  - 対応案: 互換性を確認して戻り値・引数型を明示的な型へ移行し、必要なら旧シグネチャのラッパーを残す。

- [ ] [spec] 改行エスケープ API の空エスケープ文字と未知エスケープの扱いを決める
  - 詳細: `EscapeLineSeparator(... EscSeqChar:="")` は空文字を既定の `\` として扱う一方、`UnescapeLineSeparator(... EscSeqChar:="")` は空文字をそのまま受け付け、実質的にアンエスケープしない。
  - 詳細: `UnescapeLineSeparator` は未知のエスケープ列でエスケープ文字を捨て、後続文字だけを残す。未知列を許容するのか、エラーにするのか、元の文字列を保つのかが契約として読みにくい。
  - 影響: 文字列の保存・復元で、空のエスケープ文字や未知列を含む入力だけ往復変換の結果が読み手の期待とずれる可能性がある。
  - 対応案: 空のエスケープ文字、末尾だけのエスケープ文字、`\\`、未知エスケープ列の扱いを仕様化し、`EscapeLineSeparator` / `UnescapeLineSeparator` の往復テストを追加する。

- [ ] [ref] TextFileEntity の状態プロパティ名を読み取りやすくする
  - 詳細: `AsRead` / `AsWrite` / `AsAppend` / `GetReadLock` / `GetWriteLock` は Boolean の状態プロパティだが、一般的な `Is...` / `Has...` 形式ではない。特に `GetReadLock` / `GetWriteLock` はロックを取得する操作のように読める。
  - 影響: ファイル状態を参照するだけのプロパティか、副作用を持つ操作かを名前から判断しにくい。共通テキスト入出力基盤として、利用側の読みやすさが落ちる。
  - 対応案: 互換性を保ちながら `IsReadMode`、`IsWriteMode`、`IsAppendMode`、`HasReadLock`、`HasWriteLock` などの読み取り専用プロパティを追加するか、既存名のコメントで状態値であることを明示する。

- [ ] [ref] UserInputSheet の表レイアウト前提を共通設定リーダーとして整理する
  - 詳細: `GetItemRange` は 1 列目に大項目、右隣に小項目見出し、その下に値が並ぶ表を前提にしているが、`UserInputSheet` という名前からは二段見出しの設定表専用ロジックであることが読み取りにくい。
  - 影響: 汎用の入力シート操作クラスなのか、特定レイアウトの設定テーブルリーダーなのか責務が曖昧になり、同種の読み取り処理を追加するときに置き場所を判断しづらい。
  - 対応案: `ParameterSheetReader` / `TwoLevelParameterTable` のようなレイアウトが分かる型へ切り出し、`UserInputSheet` は互換用の薄いラッパーまたは入口名にする。

## TODO 整理メモ

- `CloseWorkbook` は保存確認契約のバグ、`OpenWorkbook` は API 名とテンプレート作成挙動の仕様改善として分ける。WorkbookService 内の話でも、バグと非バグは混ぜない。
- `WorksheetService` の通常数式コピーはコピー処理のバグ、`ReadCell(GetText:=True)` は読み取り副作用と表示形式戻り値のバグ、`WorksheetService / IWorksheetService` 分割は責務整理として分ける。
- `TextFileEntity` のファイルモード、EOF、明示 Close 漏れは同じファイルライフサイクルのバグとして統合する。文字コードと改行コード指定は仕様拡張なので別 TODO のままにする。
- `FileSystemService` の実装バグ、`Lib_Common` のパス/Excel アドレス境界仕様、`FileSystemServiceTestDouble.CreateDirectory` のテストダブルキー不整合は、それぞれ実装層・仕様境界・テスト支援層の違いがあるため分ける。
- `UnitTestUtils` のキー生成はテストダブル記録基盤の整理へ統合する。ただし `CreateDirectory` の戻り値キー不整合は特定メソッドの観測可能なテスト不足なので、横断リファクタリングとは別に残す。
- `Unsign` 系 API は公開 API の綴り揺れ整理へ統合する。Doxygen コメント形式の統一はドキュメント生成規約の話なので別 TODO のままにする。
- `common_modules_repo` の同期ずれ検出は共通モジュール配布の検証、`sync.json` 対象モジュールの正本整理はプロジェクト間同期の運用設計として分ける。

## 対応を見送る事項 (無期ペンディング)

## 対応しないと決定した事項

## 対応済み事項

### 高優先度だったもの

- [x] [bug] Lib_UnitTest のテスト候補検出で複数行 Sub 宣言を見落とさない
  - 詳細: `pRunAllTest` がコードを 1 物理行ずつ正規表現にかけていたため、`_` で折り返した `Test_...` 手続き宣言を検出できなかった。
  - 最終対応: `_` による行継続を連結して論理行を作成し、その論理行を既存のテスト候補検出正規表現にかけるようにした。
  - 確認: 複数行 `Sub` 宣言の `Test_Lib_UnitTest` を追加し、結果シートに `OK` として出力されることを確認した。`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] Lib_UnitTest / UnitTestAssert でアサーション 0 件のテストを OK にしない
  - 詳細: `UnitTestAssert` はアサーション実行数を内部で数えていたが、テストランナー側は `IsFailed` だけで結果判定していたため、`Assert.*` を 1 回も呼ばないテストが OK になっていた。
  - 最終対応: `UnitTestAssert.AssertionCount` を公開し、`Lib_UnitTest.pWriteResult` で `AssertionCount <= 0` のテストを `ERR` として記録するようにした。Description には `No assertions were executed.` を出力する。
  - 確認: `Test_UnitTestAssert.bas` にアサーション件数のテストを追加し、一時プローブで 0 アサーションのテストが `ERR` になることを確認した。`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] WorksheetRangeBounds の空範囲を Shift / Transform 系で維持する
  - 詳細: 空範囲は `FinishRow = 0` や `FinishColumn = 0` で表現されるが、`Shift` で負方向へ移動すると `pFinishRow + Row` / `pFinishColumn + Column` が負値になり、`Initialize` の `pWellFormBounds` が「省略」と解釈して開始行・開始列へ補完していた。
  - 最終対応: `Shift` では開始行・開始列だけをシフトし、`FinishRow` / `FinishColumn` が `0` の場合は空方向の番兵として `0` を維持するようにした。`0` 以外の終了位置はシフト後 1 未満または上限超過をエラーにする。
  - 確認: `Test_WorksheetRangeBounds.bas` に行方向・列方向の空範囲を負方向へ `Shift` するテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] Lib_UnitTest でテスト単位の実行時エラーを ERR として記録する
  - 詳細: `pRunTestCore` は `Application.Run` の実行時エラーをテスト単位で捕捉していないため、テスト内で予期しないエラーが発生すると `UnitTestMain` 全体が中断し、そのテスト行も結果として記録されなかった。
  - 最終対応: テスト呼び出し用の一時ラッパーモジュールをランナー内部で生成し、ラッパー内の直接呼び出しで実行時エラーを捕捉するようにした。捕捉したエラーは対象テスト行へ `ERR` として記録し、`Err.Number`、`Err.Source`、`Err.Description` を Description に出す。アサート失敗の `NG`、正常終了の `OK` と状態を分けた。
  - 確認: 公開 API 追加が必要になるため恒久回帰テストは追加しない。一時プローブで実行時エラーがダイアログ停止せず `ERR` 記録になることを確認し、プローブ撤去後に `CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] UnitTestUtils の引数キー生成で Null・CVErr を安全に扱う
  - 詳細: `UnitTestUtils.pGetKeyCore` は非オブジェクト値を `pEscapeSpecial(ByVal Expression As String)` へ渡すため、`Null` や `CVErr(...)` を含む引数でキー生成自体が失敗し得た。
  - 最終対応: プリミティブ引数のキー生成を `pGetPrimitiveKey` へ分け、`CVErr(code)`、`Null`、`Empty` を文字列化より前に安定した型付きキーへ変換するようにした。
  - 確認: `Test_UnitTestUtils.bas` に `Empty`、`Null`、複数 `CVErr` の引数キーと `HasValue(Null)` のテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] UnitTestAssert で Null・CVErr・Empty を安全に比較する
  - 詳細: `pEqualsPrimitiveCore` / `pEqualsPrimitiveNumericCore` は `CStr(ExpectedValue)` / `CStr(ActualValue)` と `ExpectedValue = ActualValue` を先に実行するため、`Null` や `CVErr(...)` を含む比較で Assert 自身が型不一致や Null の不正使用で落ち得た。
  - 最終対応: Assert の表示文字列生成と比較を特殊値分岐の後に行うようにし、`CVErr` は `CLng` で取得したエラーコード、`Null` と `Empty` は各値同士だけ一致する契約にした。`EqualsNumeric` では特殊値を数値一致扱いしないまま安全に失敗させる。
  - 確認: `Test_UnitTestAssert.bas` を追加し、`Null`、`Empty`、同一/別コードの `CVErr`、`CVErr` と `Null`、`Null` と `Empty`、`EqualsNumeric(Null, Null)` のテストを追加した。`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] UnitTestAssert.EqualsArray で空配列・未初期化配列を安全に扱う
  - 詳細: `pEqualsArrayCore` は `GetArrayBounds` の戻り配列に対してすぐ `UBound` を呼ぶ。`GetArrayBounds` は未初期化配列や空配列で境界取得に失敗すると、下限・上限配列を初期化しないまま返し得た。
  - 最終対応: `EqualsArray` 側で空配列・未初期化配列を `GetArrayBounds` 前に分岐し、両側空は一致、片側だけ空は不一致として扱うようにした。あわせて `GetArrayBounds` の次元探索で発生する正常終了用の `Err.Number = 9` を戻る前に消すよう修正した。
  - 確認: `Test_UnitTestAssert.bas` に空 `Array()`、未初期化動的配列、片側空、`Null` / `CVErr` / `Empty` を含む配列のテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [spec] `IsOneRowArea` / `IsOneColumnArea` の Cell / Area 判定契約を固定する
  - 詳細: `Area` を「単一セルを除く連続した複数セル範囲」、`Cell` を「単一セル範囲」と定義するため、`IsOneRowArea` / `IsOneColumnArea` が `A1` を `False` にすること自体は仕様どおりであり、バグではなかった。一方、`Lib_Common` と `WorksheetRangeBounds` の判定名・判定粒度を揃える必要があった。
  - 最終対応: `WorksheetRangeBounds` に `IsOneRowArea` / `IsOneColumnArea` を追加し、`IsArea` は `Not IsEmpty And Not IsCell` に修正した。`Lib_Common` 側は `Range` 引数を廃止してアドレス文字列を受け取り、薄いファクトリ `New_RangeBoundsFromAddress` 経由で判定する形へ変更した。
  - 確認: `Test_Lib_Common.bas` / `Test_WorksheetRangeBounds.bas` に `A1`、`A1:B1`、`A1:A2`、`A1:B2`、`1:1`、`A:A`、空範囲、複数範囲エラーのテストを追加し、`CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] Lib_Common.IsCell の未定義関数参照を修正する
  - 詳細: `Lib_Common.IsCell` は複数範囲判定で未定義の `IsMulti(TestRange)` を呼んでいた。
  - 影響: `IsCell` を呼ぶ経路、またはプロジェクト全体のコンパイル確認で、`Sub または Function が定義されていません` になる可能性があった。
  - 最終対応: `Lib_Common` の Excel Range 判定ヘルパーを `String` 引数へ変更し、`IsCell` は `WorksheetRangeBounds.InitializeFromAddress` で生成した `WorksheetRangeBounds.IsCell` へ委譲する形にした。これにより未定義 `IsMulti` 参照は解消した。
  - 確認: `Test_Lib_Common.bas` に単一セル、複数範囲、通常範囲、行/列全体の判定テストを追加し、`CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

- [x] [req] [ref] Enumerator に読み取り専用モードを追加する
  - 詳細: `Enumerator` は `ObjectList` / `ObjectSet` などの列挙中に `Update` / `Remove` を実行できるため、呼び出し元コレクションが持つ検証や dirty 管理を迂回できる。`TimeRangeCollection` のように列挙子を読み取り用途で公開したいコレクションでは、不変条件を壊す経路になる。
  - 最終対応: `Enumerator` に `Initialize` で設定する読み取り専用状態 `IsReadOnly` を追加し、読み取り専用時は `Update` / `Remove` と `IEnumerator` 経由の更新・削除を `Class Enumerator` の明示エラーにした。`ObjectList.GetEnumerator`、`ObjectSet.GetEnumerator`、`GetArrayEnumerator` から読み取り専用指定できるようにし、`IEnumerator` 契約にも `IsReadOnly` を追加した。`WorksheetRangeBoundsEnumerator` は常に読み取り専用として返すようにした。
  - 確認: `Test_Enumerator.bas` に既定値、`Initialize` での読み取り専用設定、配列・`ObjectList`・`ObjectSet` の読み取り専用更新/削除テストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [test] ObjectList / ObjectSet / Enumerator の基礎テストを増強する
  - 詳細: `ObjectList` / `ObjectSet` / 汎用 `Enumerator` は基礎コレクションとして影響範囲が広い一方、配列変換、他コレクションからの追加、集合更新、列挙対象別の操作などが十分に固定されていなかった。
  - 影響: 基礎コレクションの修正時に、代表的な API の退行を局所的に検出しづらかった。
  - 最終対応: `ObjectList` は配列変換、`AddOther` / `AddSet`、検索開始位置と逆順検索を追加確認した。`ObjectSet` は配列変換、`AddOther` / `AddList`、`Update`、`GetContains`、`Sort` を追加確認した。`Enumerator` は ObjectList / ObjectSet 対象の `Target`、降順 `Reset`、列挙、`Update`、`Remove` を追加確認した。
  - 確認: `CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] WorksheetRangeBoundsEnumerator の SkipTo・Current 境界を Enumerator と揃える
  - 詳細: `WorksheetRangeBoundsEnumerator.SkipTo` は終端 `pLength` そのものを許す off-by-one があり、降順列挙では現在位置より小さいインデックスへ進むことを誤って禁止していた。
  - 詳細: `Current` は `MoveNext` 前や終端後の境界チェックが不足していた。`Initialize` の `EnumType` は Optional だが、省略時の空文字が `EnumerateType` 設定でエラーになっていた。
  - 影響: セル範囲走査だけが汎用 `Enumerator` と異なる境界契約になり、範囲終端や降順走査で不整合が起きる可能性があった。
  - 最終対応: `MoveNext` 失敗時に昇順は `pLength`、降順は `-1` の番兵位置へ進めるようにした。`SkipTo` は `0 To pLength - 1` の範囲と列挙方向で判定し、`Current` は有効範囲外を `Class WorksheetRangeBoundsEnumerator` の明示エラーにした。`EnumType` 省略時は Rows として初期化するようにした。
  - 確認: `Test_WorksheetRangeBoundsEnumer.bas` に `SkipTo` 境界、降順 `SkipTo`、`Current` 境界、`EnumType` 省略のテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] Enumerator の配列列挙契約を固定する
  - 詳細: 配列列挙時の `Target` が配列ではなく `pList` を返していた。
  - 詳細: 空配列の件数が現在位置に依存しており、昇順終端位置が安定していなかった。`Update` は配列でも許可する契約として固定した。
  - 影響: 配列を列挙対象にした場合だけ、コレクション列挙と異なる対象取得・更新・空判定になっていた。
  - 最終対応: 配列列挙時の `Target` は `pArray` を返すようにし、空配列の件数は安定して 0 として扱うようにした。配列での `Update` は現在要素を更新し、更新後の配列は `Target` から取得できる契約にした。`Remove` は配列では明示エラーのままとした。
  - 確認: `Test_Enumerator.bas` に配列 `Target`、空配列、配列 `Update`、配列 `Remove` のテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] Enumerator の Remove・SkipTo・Current 境界を修正する
  - 詳細: `Remove` 後に常に `pIndex = pIndex - 1` していたため、降順列挙では削除直後の次要素を飛ばしていた。
  - 詳細: `SkipTo` は終端 `Count` そのものを許す off-by-one があり、降順列挙では現在位置より小さいインデックスへ進むことを誤って禁止していた。`Current` も `MoveNext` 前や終端後の境界チェックが不足していた。
  - 影響: 列挙中削除、任意位置への移動、終端判定が不安定になり、基礎列挙子を使う全コレクション処理に波及する可能性があった。
  - 最終対応: `MoveNext` 失敗時に開始前・終端後の番兵位置へ進め、`Current` は有効範囲外を `Class Enumerator` の明示エラーにした。`Remove` 後の位置補正は昇順のみとし、`SkipTo` は有効範囲と列挙方向に応じて判定するようにした。`IEnumerator_Remove` の戻り値代入先は現行実装で正しいことを確認した。
  - 確認: `Test_Enumerator.bas` に降順削除、`SkipTo` 境界、`Current` 境界のテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] IEnumerator のインターフェイス定義を規約どおり空契約にする
  - 詳細: `IEnumerator.cls` はインターフェイスなのに `Reset` に未宣言変数を参照する実装が残っていた。
  - 詳細: `Descending` が `Private` 契約になっており、実装クラス側の公開契約と揃っていなかった。
  - 影響: インターフェイスに実装や private 契約が混ざることで、実装差し替え時のコンパイルエラーや契約読み違いにつながる可能性があった。
  - 最終対応: `Reset` を空定義にし、`Descending` を `Public` 契約へ変更した。直接生成を防ぐ `Class_Initialize` も追加し、規約どおりインターフェイス用クラスとして明示した。
  - 確認: ユニットテスト追加は不要との指示により未追加。

- [x] [bug] ObjectList の比較・存在判定・重複削除契約を統一する
  - 詳細: `ObjectList.GetIndexByItem` / `Exists` は単純な `=` 比較で、オブジェクト、`IDuplicateCheckable`、`IEquatable` の扱いが `RemoveItem` と揃っていなかった。
  - 詳細: `RemoveItem` は検索対象の型が違う場合や `Nothing` を含む場合の扱いが曖昧だった。
  - 詳細: `RemoveDuplicate` はオブジェクト要素を `Set` なしで Variant に代入し、`IEquatable` を見ていなかった。
  - 影響: 存在判定、削除、重複除去で「等しい」の意味が変わり、同じ要素を見つけられない、または誤って別要素を同一扱いする可能性があった。
  - 最終対応: `Nothing`、同一オブジェクト、`IDuplicateCheckable`、`IEquatable`、プリミティブ値を同じ比較ヘルパーで判定するようにし、`GetIndexByItem` / `Exists` / `RemoveItem` / `RemoveDuplicate` の比較契約を統一した。オブジェクト型リストへ非オブジェクトを渡した場合は `Class ObjectList` の明示エラーにした。
  - 確認: `Test_ObjectList.bas` にオブジェクト参照、`IDuplicateCheckable`、`IEquatable`、`Nothing`、型不一致のテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [spec] ObjectSet のキー生成・Nothing を契約化する
  - 詳細: `ObjectSet.pGetKey` は `Nothing` をキーにする場合に `ObjectKey = Nothing` としており、`Set ObjectKey = Nothing` ではなかった。
  - 詳細: `IDuplicateCheckable` / `IEquatable` 要素で `Nothing` を許すか、明示エラーにするかが契約化できていなかった。
  - 影響: キー生成失敗、`Nothing` の扱いが呼び出し箇所ごとに読み取りにくく、エラー原因の追跡もしづらかった。
  - 最終対応: `Nothing` は通常要素として許可し、`Set ObjectKey = ObjectItem` で Dictionary キーへ渡すようにした。キー生成の優先順位は `IDuplicateCheckable`、`IEquatable`、オブジェクト参照、プリミティブ値の順で固定した。最初に `Nothing` が入った場合も、最初の非 `Nothing` 要素で集合種別が確定する。
  - 確認: `Test_ObjectSet.bas` に `Nothing`、`IDuplicateCheckable`、`IEquatable`、キー優先順位のテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] UnitTestUtils の Dictionary 欠落キー参照を修正する
  - 詳細: `UnitTestUtils.GetValue` は `Values(pGetKey(Args))` を直接読み取っており、未設定スタブの値取得時に Dictionary を汚していた。値型の戻り値では Empty / False / 0 / 空文字相当として流れる可能性もあった。
  - 影響: 不存在取得が単なる失敗ではなく内部状態変更を伴うため、エラー捕捉後の後続処理やテストダブルの記録内容が実際とずれる可能性があった。
  - 最終対応: `GetValue` はキーを一度生成し、`Exists` で存在確認してから値を取得するようにした。不在時は `Class UnitTestUtils` の明示エラーを送出し、Dictionary の欠落キー自動作成に依存しない形へ修正した。
  - 確認: `Test_UnitTestUtils.bas` を追加し、値/オブジェクト取得と欠落キー時に Count が増えないことを確認した。既定値を持つテストダブル側は `HasValue` 分岐に寄せ、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] ObjectSet の追加・コピー・削除の実行時バグを修正する
  - 詳細: `ObjectSet.CopySet` は `result` を `New ObjectSet` せずに `result.Add` していた。
  - 詳細: `ObjectSet.Add(ItemObject, ErrorIfExists:=False)` は `pGetKey` を通さず `pDict.Exists(ItemObject)` を見ており、`IDuplicateCheckable` / `IEquatable` / オブジェクトキーと判定がずれていた。追加成功時も戻り値が `False` のままになっていた。
  - 詳細: `ObjectSet.RemoveItem` は存在しないキーを `pDict(key_var)` で読み取ってから削除していたため、`Scripting.Dictionary` の欠落キー自動作成により `IgnoreNotExists:=False` でもエラーにならず、内部状態も一時的に汚していた。
  - 影響: 集合コピー、重複無視追加、存在しない要素の削除という基本操作が失敗または誤判定し、利用側の重複排除や集合更新に波及する可能性があった。
  - 最終対応: `CopySet` の生成漏れを修正した。`Add` は `pGetKey` 後のキーで重複判定し、追加成功時に `True`、重複無視時に `False` を返すようにした。`RemoveItem` は読み取り前に `Exists` で不存在分岐し、Dictionary の欠落キー自動作成に依存しない形へ修正した。
  - 確認: `Test_ObjectSet.bas` に追加成功戻り値、`IEquatable` 重複無視追加、欠落要素削除、`CopySet` のテストを追加し、`CommonModules.xlsm` の `UnitTestMain` 全件 OK を確認済み。

- [x] [spec] ObjectList の件数 API を Count に統一する
  - 詳細: `ObjectList.Length` は廃止し、公開件数 API は `Count` に統一する。`ObjectList` 本体、`DebugInformation`、`Lib_Common.JoinStringList` / `MsgBoxPage`、`Test_ObjectList.bas` などの `Length` 参照を移行する。
  - 詳細: `Enumerator` は `Length` 優先で件数を検出しているため、`ObjectList.Length` 廃止後は `Count` のみを列挙対象の件数契約にする。
  - 詳細: `ObjectSet` は既に `Count` を持つため、`Lib_Common.JoinStringSet` の `SourceSet.Length` 参照は現状でも実行時に失敗する。
  - 影響: 件数取得の表記ゆれが残ると、`Enumerator` や共通ヘルパーが `Length` と `Count` の両方を見続ける必要があり、基礎 API の契約が曖昧になる。
  - 対応案: `Count` を追加して内部利用と共通モジュール内参照を先に移行し、テスト更新後に `Length` を削除する。`Enumerator` の件数取得を `Count` のみにし、`ObjectList`、`ObjectSet`、配列、型付きコレクションの代表ケースでテストする。
  - 最終対応: `ObjectList` の内部件数フィールドを `pCount` に変更し、範囲外エラー文言も `count` に統一した。`Test_ObjectSet.bas` / `Test_Enumerator.bas` を追加し、基本的な追加・削除・列挙・更新を固定した。追加テストで検出された `IEnumerator_Remove` の戻り値設定も修正した。

- [x] [bug] WorksheetService のセル値読み書きと空判定の契約を整理する
  - 詳細: `WriteCell(..., TypeConvert:=True)` は `CLng(Expression)` を最初に試し、成功すると即 `Exit Sub` するため、`"1.2"` や `"123.45"` のような小数文字列が `CDbl` / `CCur` に届く前に Long として丸められていた。
  - 詳細: `Expression = ""` かつ `ClearWhenEmpty:=False` の分岐は `ActiveCell`、`Select`、`Copy`、`PasteSpecial`、シート `Activate` を使って空文字列セルを作っており、選択状態とクリップボードへの副作用が大きかった。
  - 詳細: `IsEmptyCell(RangeBounds, IgnoreEmptyString:=True)` は内部ループで `pIsEmptyCellCore(cell_item)` と呼んでおり、`IgnoreEmptyString` を渡していなかった。
  - 影響: 小数・通貨値の精度喪失、非表示シートや保護シートでの書き込み失敗、`=""` 数式セルを空扱いできないことによる入力範囲終端のずれが起こり得た。
  - 最終対応: 数値文字列は `CLng` と `CDbl` の両方を試し、整数として表せる場合だけ整数、それ以外は小数として書き込むようにした。空文字列セルの作成は Excel の仕様上必要な `Copy` / `PasteSpecial` を残し、`ActiveCell` / `Select` / シート `Activate` 依存を削除した。`IsEmptyCell` は `IgnoreEmptyString` を内部へ渡し、`pIsEmptyCellCore` は `Text` ではなく `Value` の型と `Len` で判定するよう修正した。
  - 確認: `Test_WorksheetService.bas` に小数文字列、空文字列値、`=""` 数式、表示だけ空の値セルのテストを追加し、`CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] UnitTestAssert の配列比較と数値型比較を整理する
  - 詳細: `UnitTestAssert.cls`: `pEqualsMultidimensionalArrayCore` が未使用のまま残り、`ExpectedValueString` / `ActualValueString` / `CheckResult` / `actula_dimension` など未宣言変数を参照していた。
  - 詳細: `pEqualsArrayCore` は `IgnoreNumericType` を受けるが、要素比較では `pEqualsCore(..., False)` を固定指定している。`pIsNumericType` は `Currency` / `Decimal` を数値型として扱わず、VBA では通常出ない `Short` を含んでいる。
  - 影響: Assert 自体の仕様が読み取りにくく、配列比較や数値型無視比較を使うテストで誤解が生じる可能性があった。
  - 最終対応: 未使用・未完成の `pEqualsMultidimensionalArrayCore` を削除した。配列比較は現行仕様どおり型も含めた厳密比較とし、`IgnoreNumericType` にかかわらず要素比較では数値型差を無視しない旨を内部コメントに追記した。`Currency` / `Decimal` は型無視の数値比較対象に含めない旨を `EqualsNumeric` / `NotEqualsNumeric` の詳細コメントへ追記した。
  - 確認: `CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] WorksheetService のコピー・書式操作 API を契約どおりにする
  - 詳細: `pCopyCellCore(..., CopyNumberFormat, ...)` は `CopyNumberFormat` を受け取っているが、値コピー、通常数式、配列数式、空セルの各分岐で常に `NumberFormatLocal` をコピーしていた。
  - 詳細: `CopyRange` は `pCalcurateArrayCount(arr_rows, arr_cols, ...)` で配列数式の貼り付けサイズを計算しているが、`arr_rows` / `arr_cols` をその後使っていなかった。計算自体も終端セルを含む個数として off-by-one になる可能性があった。
  - 詳細: `SetAlignment` の `Case xlRight` で `xlRightr` を代入し、縦位置 `Case Else` では `target_range.HorizontalAlignment = xlBottom` としていた。
  - 詳細: `ClearRange(..., ClearColors:=True, ClearNumberFormats:=False)` でも `pClearRangeCore` が `TargetRange.NumberFormat = "General"` を実行していた。`SetSheetTabColor` / `SetRangeColor` は `RGB(0, 0, 0)` / `vbBlack` のような 0 の色値を `Color` として指定できなかった。
  - 最終対応: `CopyNumberFormat` が True の場合だけ表示形式をコピーするよう修正した。配列数式コピーは計算済みサイズを使い、終端セルを含む個数へ修正した。`SetAlignment` の typo と縦位置代入先、`ClearColors` 時の表示形式クリア、0 以上の色値指定を修正した。
  - 確認: `Test_WorksheetService.bas` に再発防止テストを追加し、`CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

### 中優先度だったもの

- [x] [spec] Excel Range 判定ヘルパーの名前と判定粒度を整理する
  - 詳細: `Lib_Common` の `IsMultiRange`、`IsArea`、`IsCell`、`IsEntireRow`、`IsEntireColumn`、`IsOneRowArea`、`IsOneColumnArea` は Excel `Range` を直接受けていたが、`WorksheetRangeBounds` 側にも似た判定があり、判定基準のずれが起きやすかった。
  - 影響: 呼び出し側が Excel `Range` と `WorksheetRangeBounds` のどちらの判定を使うべきか迷いやすく、セル 1 個、空範囲、行/列全体、複数選択の扱いを誤解しやすかった。
  - 最終対応: `Lib_Common` の判定ヘルパーはアドレス文字列引数へ破壊的変更し、単一矩形範囲の判定は `WorksheetRangeBounds` へ委譲する形にした。`SplitExcelAddress` を厳密化し、`SplitA1RangeAddress`、`WorksheetRangeBounds.InitializeFromAddress`、薄いファクトリ `New_RangeBoundsFromAddress` を追加した。複数範囲は `WorksheetRangeBounds` で表現できないため、`IsMultiRange` 以外では明示エラーにした。
  - 確認: `Test_Lib_Common.bas` / `Test_WorksheetRangeBounds.bas` にアドレス分解、文字列判定、複数範囲エラー、空範囲 `IsArea=False` のテストを追加し、`CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] WorkbookService の保存・シート操作・Excel 状態復元をまとめて修正する
  - 詳細: `pGetTargetSheetIndex` の `SheetIndex < 0` 分岐が `TargetBook.Worksheets.Count + 1 - SheetIndex` になっており、`SheetIndex:=-1` が末尾ではなく `Count + 2` になっていた。
  - 詳細: `SaveWorkbook` は `On Error Resume Next` で `SaveAs` を実行し、失敗時に `Debug.Print` するだけだった。`GetExcelFileFormat` は拡張子判定が大小文字に依存し、`.XLSM` や `.Csv` が既定の `xlOpenXMLWorkbook` へ落ちていた。
  - 詳細: `OpenWorkbook`、`SaveWorkbook`、`RemoveWorksheet`、`CopyWorksheet`、`Lib_UnitTest.UnitTestMain` は Excel 設定やアクティブブックを個別に退避復元しており、例外時の復元経路が不足していた。
  - 最終対応: 負数 `SheetIndex` は `-1` を末尾として扱う式へ修正した。`GetExcelFileFormat` は `LCase$` で拡張子の大小文字を無視するようにした。`SaveWorkbook` は `SaveAs` 失敗を再送出し、保存対象ウィンドウ、アクティブブック、Excel 設定を成功時・失敗時とも復元するようにした。`OpenWorkbook`、`RemoveWorksheet`、`CopyWorksheet`、`UnitTestMain` も `ApplicationScreenUpdateManager` とエラー経路で復元する形へ整理した。
  - 確認: `Test_WorkbookService.bas` と `Test_Lib_Common.bas` に負数 `SheetIndex`、保存失敗、拡張子大小文字のテストを追加し、`CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

- [x] [bug] WorksheetService の UsedRange 取得と範囲クリア境界を修正する
  - 詳細: `pGetRawUsedRange` の空領域判定が `If used_s_row = used_s_row And used_s_col = used_f_col ...` となっており、開始行と終了行の比較が常に True になっていた。`used_s_row = used_f_row` の誤記だった。
  - 詳細: `pGetFirstRowCore` / `pGetFirstColumnCore` は `srch_rng.Cells(srch_rng.Cells.Count)` を使うため、全シート級の範囲では Excel の `Range.Count` 桁あふれが発生していた。
  - 詳細: `ClearRange` は対象範囲をクリアしたあと、非表示行を補正するループで `RangeBounds.Row To used_row`、非表示列を補正するループで `RangeBounds.Column To used_col` を走査しており、対象範囲外の非表示行・列まで触り得た。
  - 最終対応: UsedRange の 1 セル判定を `used_s_row = used_f_row` に修正した。最初の行・列検索では `srch_rng.Cells(srch_rng.Rows.Count, srch_rng.Columns.Count)` を使い、`Range.Count` への依存を削除した。`ClearRange` の非表示行・列補正は `RangeBounds.FinishRow` / `RangeBounds.FinishColumn` を上限に制限した。
  - 確認: `Test_WorksheetService.bas` に 1 列 UsedRange、全シート級範囲、対象範囲外の非表示行・非表示列を突くテストを追加し、`CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

### 低優先度だったもの

- [x] [ref] インターフェイス定義のタグと直接生成防止を規約へ揃える
  - 詳細: `IWorkbookService.cls`、`IFileSystemService.cls`、`ITextFileService.cls`、`ITextFileEntity.cls` のタグが `'# Interface` で、規約の `'#Interface` と揺れていた。インターフェイス直接生成防止の `Class_Initialize` も横断確認対象だった。
  - 最終対応: 4 インターフェイスのタグを `'#Interface` に統一した。全インターフェイスで `Class_Initialize` が規約文言のエラーを送出することを確認した。
  - 確認: `CommonModules.xlsm` へ import 後に `UnitTestMain` 全件 OK を確認済み。

- [x] [ref] 手動検証用モジュールを整理する
  - 詳細: `CommonModules/modules/Module1.bas` は `Test_ApplicationScreenUpdateManager.bas` とほぼ同じ内容で、`Tmp_ManualTest.bas` は `Option Explicit` がなく、個人環境の絶対パスや `Debug.Print` を含む。
  - 最終対応: Module1.bas を削除した。
