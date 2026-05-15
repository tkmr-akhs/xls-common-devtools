# 共通モジュール TODO

このファイルは、現時点で見つかっている未対応事項や設計上の検討事項をまとめたものです。

## タグ一覧

| タグ名 | 短縮タグ | 意味 |
| --- | --- | --- |
| バグ | `[bug]` | 現仕様に対する誤動作。実務データで誤った結果になるもの。 |
| 要リファクタリング | `[ref]` | 動いてはいるが、設計・責務・依存関係・テスト容易性に問題があるもの。 |
| 仕様改善 | `[spec]` | 仕様の明確化、境界条件の整理、仕様上の矛盾、現仕様として明確だが著しく不便なものの見直し。 |
| 必須機能不足 | `[req]` | 実務利用に必要だが、現時点で仕様として欠けているもの。 |
| 利便性向上 | `[ux]` | 結果の見やすさ、調査しやすさ、操作性、運用しやすさの改善。 |
| 製品試験増強 | `[test]` | 製品試験の考慮不足など。 |

## 高優先度

- [ ] [bug] WorksheetService.InsertRows / DeleteRows で行範囲以外を拒否する
  - 詳細: `InsertRows` / `DeleteRows` は `RangeBounds.Row` から `RangeBounds.FinishRow` だけを使って `"行:行"` の Range を作るため、列範囲や通常のセル範囲を渡しても行全体の挿入・削除として実行される。特に `A:A` のような列範囲では全行相当を対象にし得る。
  - 影響: 呼び出し側が範囲オブジェクトを誤って渡したとき、明示エラーではなくシート構造を大きく変更する可能性がある。基盤 API として、範囲単位操作と行単位操作の境界が危険に曖昧になる。
  - 対応案: `RangeBounds.IsEntireRow` または行方向の範囲だけを受け付ける契約に固定し、列範囲、セル範囲、複数行範囲、空範囲のテストを追加する。

- [ ] [bug] ObjectList / ObjectSet の検索・削除 API で要素型を検査する
  - 詳細: `ObjectList.Exists` / `GetIndexByItem` / `RemoveItem` は特殊値チェックだけで `pCheckItemType` を通さず、プリミティブ比較では `ItemObject1 = ItemObject2` の暗黙変換に任せている。そのため数値 `1` のリストに文字列 `"1"` を渡すと同一扱いになり得る。
  - 詳細: `ObjectSet.Exists` / `GetContains` / `RemoveItem` も追加時と同じ型検査を通さず、`IDuplicateCheckable` / `IEquatable` のキー生成経路では型違いの値が実装詳細由来の実行時エラーになる可能性がある。
  - 影響: 追加・更新では型を固定しているのに、検索・削除では型違いを False、暗黙一致、実行時エラーのどれにするかが揃わない。コレクション基盤として「同じ要素」の意味が操作ごとに変わる。
  - 対応案: 検索・削除系 API でも保存済み型との照合を行い、型違いは明示エラーまたは常に不一致のどちらかに統一する。数値と数値文字列、Boolean と数値、IEquatable 実装と非オブジェクト引数のテストを追加する。

- [ ] [bug] WorksheetService.SetSheetOutlineLevel の既定値 0 を Excel に渡さない
  - 詳細: `SetSheetOutlineLevel` は `RowLevels:=0`、`ColumnLevels:=0` を既定値にし、そのまま `target_sheet.Outline.ShowLevels(RowLevels:=RowLevels, ColumnLevels:=ColumnLevels)` へ渡している。Excel のアウトラインレベルは 1 以上の指定または引数省略が前提で、0 を渡すと実行時エラーになり得る。
  - 影響: 引数省略で「変更しない」つもりの呼び出しや、行だけ・列だけを指定する呼び出しが、アウトライン操作ではなく基盤 API 側の既定値で失敗する。
  - 対応案: `0` は未指定を表す値として扱い、行・列それぞれ 1 以上のときだけ `ShowLevels` へ渡す。`0`、行のみ、列のみ、両方指定、範囲外値のテストを追加する。

- [ ] [bug] WorksheetService.Sort のキー列番号と並び順を事前検証する
  - 詳細: `WorksheetService.Sort` は `SortKeyAndOrder` の個数が偶数であることだけを確認し、キー列番号が `RangeBounds` の列数内か、並び順が `xlAscending` / `xlDescending` かを検証していない。`0`、負数、範囲外列、無効な並び順が Excel 呼び出しへそのまま渡る。
  - 影響: 呼び出し側の指定ミスで、対象範囲外の列をキーにしたソートや Excel 由来の実行時エラーが発生する。データ並べ替えの基盤 API として、誤ったデータ移動の原因を特定しにくい。
  - 対応案: キー列番号は `1 To RangeBounds.ColumnCount`、並び順は未指定時の既定値または Excel の有効値だけに制限する。`0`、負数、列数超過、無効な並び順のテストを追加する。

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

- [ ] [bug] WorksheetService.ReadCell(GetText:=True) の副作用と表示形式戻り値を修正する
  - 詳細: `ReadCell(GetText:=True)` は `Range.Text` を読むために列幅を変更するが、元の `ColumnWidth` を `Long` に保存しているため、小数を含む列幅を正確に復元できない。途中でエラーが発生した場合に列幅が戻らない経路もある。
  - 詳細: `GetText:=False` では `NumberFormatLocal` を返す一方、`GetText:=True` では `NumberFormat` を返しており、インターフェイスコメントと `WriteCell` の表示形式指定と揃っていない。
  - 影響: 読み取り API の呼び出しだけでシートの列幅が変わる可能性があり、表示形式の戻り値もロケール依存の読み書きで不整合になる。
  - 対応案: 列幅は `Double` で退避し、エラー時も復元する。`NumberFormat` の返却値は `NumberFormatLocal` に統一し、列幅保持とローカル表示形式のテストを追加する。

- [ ] [bug] WorksheetService.ActivateRange で対象シートを確実にアクティブ化する
  - 詳細: `ActivateRange` は `target_sheet.Range(...).Activate` を直接呼ぶだけで、対象ブックや対象シートを先にアクティブにしていない。
  - 影響: 対象シートがアクティブでない状態では `Range.Activate` が失敗したり、呼び出し側が期待した範囲が選択されない可能性がある。
  - 対応案: UI 操作用 API として対象ブック、対象シート、対象範囲のアクティブ化順序を明示し、非アクティブブック/非アクティブシートからの呼び出しをテストする。

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

- [ ] [bug] FileSystemService.GetLastModified で存在しないパスを明示エラーにする
  - 詳細: `GetLastModified` は `IsFile(path_str)` が False の場合に `pFSO.GetFolder(path_str)` へ進むため、存在しないパスやファイル/フォルダー以外の指定が、共通基盤の明示エラーではなく FSO の実行時エラーになる。
  - 影響: 更新日時比較や最新ファイル選択の前処理で、対象不存在とフォルダー取得失敗の区別がつかない。テストダブル側も未登録値は `UnitTestUtils.GetValue` のキー不一致エラーで、実装と同じ契約を固定しにくい。
  - 対応案: `PathExists` / `IsFile` / `IsDirectory` を使って、存在しないパス、ファイル、フォルダーを入口で分岐し、存在しない場合は `Source:="Class FileSystemService"` の明示エラーにする。存在しないファイル、存在しないフォルダー、正常ファイル、正常フォルダーのテストを追加する。

- [ ] [bug] TextFileEntityTestDouble の ReadLine / WriteLine でオープン状態とモードを検査する
  - 詳細: `ReadLine` は未オープンや書き込みモードでも、登録済みの `ReadLine_Values(ReadCount)` があれば値を返す。`WriteLine` も未オープンや読み取りモードを確認せず、常に `WriteLine_Results` へ記録する。
  - 影響: 実装の `TextFileEntity` では未オープン、読み書きモード違反がエラーになるのに、テストダブルでは成功するため、呼び出し側の Open/Close 順序やモード指定ミスをテストで見逃す。
  - 対応案: `pCheckOpened` 相当と読み取り/書き込みモード検査をテストダブルにも入れる。未オープン Read/Write、読み取りモード Write、書き込みモード Read、Close 後の Read/Write をテストに追加する。

- [ ] [bug] WorkbookServiceTestDouble.RemoveVBComponents の配列キー生成を衝突しない形にする
  - 詳細: `RemoveVBComponents` は `pBuildComponentNamesKey(ComponentNames)` で配列をカンマ連結した文字列にしてから `UnitTestUtils` のキーへ渡している。配列要素内のカンマをエスケープせず、要素型や配列境界も含めないため、`Array("A,B", "C")` と `Array("A", "B,C")` のような指定を区別できない。
  - 影響: VBComponent 削除のスパイ記録で別の呼び出しが同じキーになり、テストが誤った呼び出しを検出または見逃す可能性がある。`Null` や `CVErr(...)` を含む配列では `CStr` 変換時の実行時エラーにもなる。
  - 対応案: 配列引数は要素数、境界、型タグ、エスケープ済み値を含めてキー化するか、`UnitTestUtils` 側の配列キー対応へ寄せる。カンマを含むモジュール名、単一要素配列とスカラー、特殊値配列のテストを追加する。

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

- [ ] [bug] FileSystemService の Force 上書きでコピー元・移動元を削除しない
  - 詳細: `MoveDirectory` / `CopyDirectory` / `MoveFile` / `CopyFile` は `Force:=True` のとき、実操作前に `pRemoveOldPath(src_path, dst_path, ends_with_sep)` で既存のコピー先・移動先を削除する。`SourcePath` と `DestinationPath` が同じ実体を指す場合や、末尾セパレータ付きの `DestinationPath` 配下に元の名前を置く指定がコピー元・移動元自身を指す場合、実操作前に元データを削除し得る。
  - 影響: 上書き指定のつもりで同一パス、親ディレクトリ、相対パス表記違いを渡すと、コピー / 移動が失敗したうえで元ファイルや元ディレクトリだけが消える可能性がある。
  - 対応案: 正規化後の実コピー先・実移動先を先に確定し、コピー元・移動元と同一またはその配下 / 親子関係で危険な場合は削除前に明示エラーにする。同一ファイル、同一ディレクトリ、親ディレクトリ末尾セパレータ、相対パスと絶対パスの表記違いのテストを追加する。

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

- [ ] [bug] Enumerator.Update の配列対象更新を呼び出し元へ反映しないまま成功扱いにしない
  - 詳細: `Enumerator.Initialize` は `TargetCollection As Variant` を値渡しで受け取り、配列の場合は `pArray = NewValue` と内部フィールドへコピーする。`Update` は `pArray(pIndex)` を更新するだけなので、`GetArrayEnumerator(source_arr).Update(...)` を呼んでも元の `source_arr` は変わらない。
  - 影響: ObjectList / ObjectSet の列挙子では `Update` が対象コレクションを更新するのに、配列列挙子だけは見かけ上成功して内部コピーだけが変わる。共通列挙インターフェイスとして書き込み可否が対象種別で予測しにくい。
  - 対応案: 配列列挙子は常に読み取り専用にするか、ByRef で配列を保持できる別 API に分ける。配列、ObjectList、ObjectSet それぞれの `Update` / `Remove` 契約テストを追加する。

- [ ] [bug] Lib_Common.IsQuotedWith で空の引用文字列を一致扱いにしない
  - 詳細: `IsQuotedWith(Expression, QuoteString:="")` は `EndString` も空文字に補完され、`Left(Expression, 0)` / `Right(Expression, 0)` がどちらも空文字になるため、長さ 2 以上の任意文字列を引用済みとして True にする。
  - 影響: 引用符設定が未指定または設定ミスで空文字になった場合、実際には囲われていない値を囲われているものとして扱う。文字列の前後トリム、クォート除去、設定値パースなどの前提判定に使うと後続処理を誤る可能性がある。
  - 対応案: `QuoteString` と補完後の `EndString` が空文字の場合は False または明示エラーにする。空引用、片側だけ空、通常の同一引用符、開始・終了が異なる引用符のテストを追加する。

- [ ] [spec] WorkbookService.OpenWorkbook の名称と実挙動を揃える
  - 詳細: `OpenWorkbook` は `Workbooks.Open` ではなく `Workbooks.Add(Template:=file_path)` でテンプレートから非表示ブックを作成している。コメントには元ファイルをロックしない旨があるが、API 名からは通常の open と誤解しやすい。
  - 影響: 呼び出し側が既存ブックを開いて操作する API と誤認し、保存先やロック、元ファイル更新の前提を取り違える可能性がある。
  - 対応案: 互換性を考慮しつつ `CreateWorkbookFromTemplate` / `OpenWorkbookCopy` など実態に近い名前を追加し、旧名はラッパーまたは非推奨コメントにする。

- [ ] [spec] TextFileEntity / TextFileService で文字コードと改行コードを指定できるようにする
  - 詳細: 現状は VBA の `Open ... For Input/Output/Append` と `Print #` に依存しており、UTF-8 BOM、Shift_JIS、CRLF/LF などを呼び出し側が明示できない。
  - 影響: 共通モジュールは UTF-8 BOM の `.md` / `.ps1` と Shift_JIS の `.bas` / `.cls` を扱うため、基盤テキスト API として文字コードの契約がないと使える範囲が狭い。
  - 対応案: 既存 API は互換維持し、必要に応じて encoding / newline を指定できるオプションまたは別サービスを追加する。代表的な文字コードと末尾改行のテストを追加する。

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

- [ ] [ux] Lib_UnitTest の再実行ボタンは表示行のテストを実行する
  - 詳細: 再実行ボタンは作成時点の行番号を `Name` に持ち、`UnitTestMain` は `Application.Caller` を数値化して、その行の `Category` / `Test Item` を再実行している。通常のフィルターや並べ替えだけでは直ちにずれないが、行挿入・行削除などでボタンの現在位置と `Name` の行番号がずれると、見た目の行とは別のテストを再実行し得る。
  - 影響: 結果シートを手動編集した後の再実行で、意図と違う行の結果を更新する可能性がある。テスト実行自体の正否より、結果確認時の操作性・誤操作耐性の問題として扱う。
  - 対応案: 再実行対象はボタン名の行番号ではなく、`Application.Caller` で得た Shape の `TopLeftCell.Row` から現在行を取る、またはボタンにモジュール名・テスト名を保持して行移動に依存しない形へ変更する。行挿入・行削除後の再実行確認を追加する。

- [ ] [spec] WorksheetService の書式設定 API で省略引数を変更なしとして扱えるようにする
  - 詳細: `SetRangeColor` は背景色だけを指定したい呼び出しでも既定値の `FontColorIndex:=0` を書き込み、`SetAlignment` は `Orientation` や `IndentLevel` だけを変えたい呼び出しでも水平・垂直配置を既定値へ設定する。省略した項目を「変更しない」と扱う API と、既定値へリセットする API が分かれていない。
  - 影響: 共通の書式設定 API として、呼び出し側が一部の書式だけを変更したい場合に既存書式を壊しやすい。既存値を保つには呼び出し側が事前に読み取って全引数を渡す必要があり、基盤として使いづらい。
  - 対応案: 互換性を保つ場合は既存 API のリセット挙動を明記し、`SetRangeFontColor` / `SetRangeInteriorColor` / `SetHorizontalAlignment` など単一責務 API を追加する。破壊的変更を許す場合は Optional 引数の未指定を判別できる `Variant` や別オプションで「変更なし」を表現する。

- [ ] [spec] WorksheetService.ClearRange の ClearAll と Excel Clear の関係を明確にする
  - 詳細: `ClearRange(ClearAll:=True)` は内容、表示形式、文字色・背景色、コメント、ハイパーリンクを個別に消すが、罫線、入力規則、条件付き書式、スタイルなどは `Range.Clear` のようには消さない。一方で `IWorksheetService` のコメントは「すべてクリア」と読める。
  - 影響: 呼び出し側が Excel の「すべてクリア」と同等の初期化を期待すると、テンプレート由来の罫線や入力規則などが残り、後続処理や見た目に意図しない状態が残る。逆に現行の限定クリアが意図なら、名前と契約から対象外の書式を判断しづらい。
  - 対応案: `ClearAll` を Excel `Range.Clear` 相当にするか、現行挙動を維持するなら「個別オプション全指定」としてコメント・名称を整理する。罫線、入力規則、条件付き書式、ハイパーリンク、コメントを含むテストで契約を固定する。

- [ ] [ux] WorksheetService.XLookup の検索範囲と戻り範囲の形状を検証する
  - 詳細: `XLookup` は `LookupRangeBounds` と `ReturnRangeBounds` をそのまま `=XLOOKUP(...)` の式に埋め込み、検索範囲が 1 行または 1 列か、戻り範囲の行列数が検索方向と整合するかを確認していない。
  - 影響: 2 次元検索範囲や行数・列数が合わない戻り範囲を渡すと、基盤 API の引数エラーではなく Excel の評価結果 `#VALUE!` などとして返り、データ未検出・式エラー・仕様違反の区別がつきにくい。
  - 対応案: 検索範囲は 1 行または 1 列に制限し、戻り範囲は検索方向の長さを一致させる。`MatchMode` / `SearchMode` も Excel の有効値だけを受け付けるテストを追加する。

- [ ] [ux] WorksheetRangeBoundsEnumerator.Initialize で Nothing の列挙対象を拒否する
  - 詳細: `Initialize(TargetCollection As WorksheetRangeBounds)` は `TargetCollection Is Nothing` を確認せず、`Target` に `Nothing` を設定したまま初期化済みにできる。`pCalculateLength` は `pRangeBounds Is Nothing` で何もせず、列挙子は長さ 0 のように振る舞う。
  - 影響: 呼び出し側の範囲生成漏れが初期化時に検出されず、`MoveNext` が False を返すだけの空列挙として扱われる。`Target` や `Current` で後から落ちる場合もあり、基盤列挙子として原因を追いにくい。
  - 対応案: `Initialize` の入口で `Nothing` を `Class WorksheetRangeBoundsEnumerator` の明示エラーにする。`Nothing`、空範囲、通常範囲の初期化テストを追加する。

- [ ] [ux] Enumerator.Initialize で Nothing の列挙対象を明示的に拒否する
  - 詳細: `Enumerator.Initialize(TargetCollection As Variant, ...)` はオブジェクト引数が `Nothing` でも初期化済みにできる。後続の `Reset` / `MoveNext` / `Current` では `pList.Count` や `pList.Item(...)` で汎用的なオブジェクト未設定エラーになり、配列でない値を渡したときの `Class Enumerator` の明示エラーと扱いが揃っていない。
  - 影響: `ObjectList` / `ObjectSet` の生成漏れやテストダブルの戻り値設定漏れが、列挙子初期化時に分からず、後続操作の場所で原因を追う必要がある。
  - 対応案: `Initialize` の入口で `Nothing` を `Class Enumerator` の明示エラーにする。`Nothing`、空の `ObjectList` / `ObjectSet`、空配列、通常配列の初期化テストを追加する。

- [ ] [ux] WorksheetRangeBounds.Intersect で Nothing を明示的に拒否する
  - 詳細: `Intersect(ByVal OtherRangeBounds As WorksheetRangeBounds)` は `OtherRangeBounds Is Nothing` を確認せず、すぐに `OtherRangeBounds.WorksheetName` / `WorkbookName` / `Row` などを参照する。
  - 影響: 呼び出し側の範囲生成漏れが `Class WorksheetRangeBounds` の説明付きエラーではなく、VBA の汎用的な「オブジェクト変数または With ブロック変数が設定されていません。」として表面化する。範囲交差を使う `WorksheetService.GetUsedRangeBounds` などでも原因を追いにくい。
  - 対応案: 入口で `Nothing` を明示エラーにし、`Nothing`、未初期化範囲、空範囲、非交差範囲、正常交差範囲のテストを分けて追加する。

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

- [ ] [ref] UnitTestAssert のコメントアウト済み旧配列比較コードを削除する
  - 詳細: `UnitTestAssert.cls` には、未使用の `pEqualsMultidimensionalArrayCore` がコメントアウトされたまま残っている。対応済み事項では削除済みと記録されているため、現在のソースと TODO 履歴の説明も食い違っている。
  - 影響: 配列比較の現行仕様と、復活予定があるのか不明な旧実装案が同じ場所に並び、テスト失敗時や拡張時に読むべき実装を判断しにくい。
  - 対応案: コメントアウト済みコードは削除し、必要な設計メモだけを TODO またはコメントへ短く残す。多次元配列の差分表示改善は既存 TODO の `EqualsArray` 表示改善へ集約する。

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

- [ ] [ux] Lib_UnitTest のテスト実行順を安定化する
  - 詳細: `pRunAllTest` は `VBProject.VBComponents` の列挙順と各 `CodeModule` の出現順をそのまま結果シートに出力する。モジュールの import/export や VBIDE の内部順序変更で同じテスト集合でも行順が変わり得る。
  - 影響: 前回結果との目視比較、特定行の再実行ボタン、CI 風のログ比較で差分が出やすく、失敗したテストの追跡がしづらい。
  - 対応案: `Category`、`Test Item` の昇順など、実行順・表示順を仕様化する。依存順が必要なテストは禁止または別タグ化する。

- [ ] [ref] UnitTestAssert のオブジェクト比較を参照同一性と値同一性に分ける
  - 詳細: `Equals` / `NotEquals` のオブジェクト比較は `Is` による参照同一性だけで、`IEquatable` を実装する値オブジェクトでも内容比較は行わない。一方、`ObjectList` / `ObjectSet` や `UnitTestUtils` は `IEquatable.GetIdentityString` を値同一性として使っている。
  - 影響: `WorksheetRangeBounds` などの値オブジェクトをテストするとき、同じ範囲を表す別インスタンスは `Equals` で失敗し、呼び出し側が毎回 `GetIdentityString` や個別プロパティ比較へ落とす必要がある。
  - 対応案: 既存 `Equals` は参照同一性として維持しつつ、`EqualsObject` / `EqualsValueObject` / `SameReference` など名前で意図が分かる API を追加するか、`IEquatable` 対応の方針を明記する。

- [ ] [ux] Lib_UnitTest の結果区分を失敗種別ごとに分ける
  - 詳細: 現在は assertion 未実行、テスト本体の実行時エラー、実行用一時モジュール側の runner error がすべて `ERR` になる。説明文を読めば区別できるが、Result 列だけでテスト失敗、テスト定義不備、テスト基盤エラーを絞り込めない。
  - 対応案: `NO_ASSERT`、`RUNTIME_ERR`、`RUNNER_ERR` のような内部区分列を追加するか、Result 列の値を分ける。既存の `OK` / `NG` / `ERR` フィルター運用と互換性を取る。

- [ ] [ref] ObjectList / ObjectSet の相互取り込みメソッド名を型名付きへ揃える
  - 詳細: `ObjectList.AddOther` は別の `ObjectList` を取り込み、`ObjectSet.AddOther` は別の `ObjectSet` を取り込む。同時に `ObjectList.AddSet` / `ObjectSet.AddList` も存在するため、`AddOther` だけでは取り込み元の型が名前から分かりにくい。
  - 対応案: 既存名は互換ラッパーとして残し、`AddList` / `AddSet` または `AddRangeFromList` / `AddRangeFromSet` のように型と複数追加が分かる名前へ寄せる。公開コメントでも重複時・型違い時の挙動を明示する。

- [ ] [ux] TextFileEntity の入出力状態エラー番号を独自エラー番号に統一する
  - 詳細: `OpenFile` のオープン済みチェック、`AsAppend And Not AsWrite`、`ReadLine` の書き込みモード拒否、`WriteLine` の読み取りモード拒否で `Err.Raise vbObjectError` を使っている。他の共通クラスの多くは `vbObjectError + 1` を使っており、独自エラー番号の扱いが揃っていない。
  - 影響: テキストファイル基盤の状態エラーだけが VBA/COM の基底値そのものになり、呼び出し側やテストが共通基盤の明示エラーを番号で識別しにくい。`ErrorRaised` で具体的な番号を確認するテストも書きづらい。
  - 対応案: `Err.Raise Number:=vbObjectError + 1, Source:=..., Description:=...` 形式へ揃える。オープン済み、Append 指定矛盾、読み取り/書き込みモード違反のエラー番号・Source・Description をテストで固定する。

- [ ] [ux] WorksheetRangeBounds の範囲外エラー番号を独自エラー番号に統一する
  - 詳細: `Item` / `pGetCellVertical` / `pGetCellHorizontal` の範囲外・空範囲エラーは `Err.Raise vbObjectError` を使っており、同じクラス内の他メソッドで使っている `vbObjectError + 1` と異なる。
  - 影響: 呼び出し側やテストが共通基盤の明示エラーを番号で識別する場合、`WorksheetRangeBounds` の一部境界エラーだけ別番号になり、同じ入力検証失敗として扱いにくい。エラー番号が VBA/COM の基底値そのものになるため、診断上も意図が読み取りにくい。
  - 対応案: `Err.Raise Number:=vbObjectError + 1, Source:=..., Description:=...` 形式へ揃える。`Item(-1)`、空範囲の `Item(0)`、範囲外 `Item` のエラー番号・Source・Description をテストで固定する。

- [ ] [ref] 比較・同一性・重複判定インターフェイスの責務名を整理する
  - 詳細: `IComparable.IsEqualTo` は整列上の同順位、`IEquatable.Equals` は同一性、`IDuplicateCheckable.IsDuplicateOf` / `GetKey` は集合内の重複判定を表している。いずれも「等しい」に近い概念だが、名前と優先順位が分散しており、`ObjectList` / `ObjectSet` では `IDuplicateCheckable` が `IEquatable` より優先される。
  - 影響: 値オブジェクトを追加する利用者が、同一性・同等性・重複キー・ソート同順位のどれを実装すべきか判断しづらい。基盤コレクションに渡したときの比較基準も名前から読み取りにくい。
  - 対応案: 既存互換を維持しつつ、`IIdentityComparable`、`ISetKeyProvider`、`IOrderComparable` など責務が分かる名前やガイドコメントへ整理する。`ObjectList` / `ObjectSet` 側の比較優先順位も公開コメントに明記する。

- [ ] [ux] Dictionary 由来の重複キー・範囲外エラーを共通基盤のエラーに包む
  - 詳細: `CounterSet.AddCounter`、`ObjectSet.Add`、`ObjectSet.Item` / `Update` / `Remove` などは、重複キーや範囲外インデックスを `Scripting.Dictionary` の実行時エラーに任せる経路がある。共通クラス自身が投げる `Err.Raise vbObjectError + 1` と、実装依存のエラー番号・文言が混在している。
  - 影響: 呼び出し側やテストが失敗理由を `Source` / `Description` で追うとき、基盤クラスの契約違反なのか内部 Dictionary の例外なのかが読み取りにくい。エラー番号を基盤独自に揃える作業は低優先度の利便性改善として扱う。
  - 対応案: 公開 API の入口で重複・存在・インデックス範囲を検証し、`Source:="Class ..."` の独自エラーへ変換する。重複追加、空集合アクセス、範囲外更新、存在しない削除のエラー番号・Source・Description をテストで固定する。

## TODO 整理メモ

- `CloseWorkbook` は保存確認契約のバグ、`OpenWorkbook` は API 名とテンプレート作成挙動の仕様改善として分ける。WorkbookService 内の話でも、バグと非バグは混ぜない。
- `WorksheetService` の通常数式コピーはコピー処理のバグ、`ReadCell(GetText:=True)` は読み取り副作用と表示形式戻り値のバグ、`WorksheetService / IWorksheetService` 分割は責務整理として分ける。
- `TextFileEntity` のファイルモード、EOF、明示 Close 漏れは同じファイルライフサイクルのバグとして統合する。文字コードと改行コード指定は仕様拡張なので別 TODO のままにする。
- `FileSystemService` の実装バグ、`Lib_Common` のパス/Excel アドレス境界仕様、`FileSystemServiceTestDouble.CreateDirectory` のテストダブルキー不整合は、それぞれ実装層・仕様境界・テスト支援層の違いがあるため分ける。
- `UnitTestUtils` のキー生成はテストダブル記録基盤の整理へ統合する。ただし `CreateDirectory` の戻り値キー不整合は特定メソッドの観測可能なテスト不足なので、横断リファクタリングとは別に残す。
- `Unsign` 系 API は公開 API の綴り揺れ整理へ統合する。Doxygen コメント形式の統一はドキュメント生成規約の話なので別 TODO のままにする。
- `common_modules_repo` の同期ずれ検出は共通モジュール配布の検証、`sync.json` 対象モジュールの正本整理はプロジェクト間同期の運用設計として分ける。

## 対応を見送る事項 (無期ペンディング)

- [ ] [bug] UnitTestUtils の配列引数と区切り文字入り同一性キーを安全に扱う
  - 詳細: `UnitTestUtils.pGetKeyCore` は非オブジェクト引数を `pGetPrimitiveKey` へ渡し、最終的に `CStr(ArgItem)` するため、テストダブル対象メソッドの引数を配列のままキー化すると型不一致になる。
  - 詳細: `IEquatable` 引数では `GetIdentityString()` をそのまま `|` 連結に使い、プリミティブ値のような `|` / `<` / `>` のエスケープを行っていない。実装クラスの同一性文字列に区切り文字が入ると、複数引数キーの衝突や誤読が起こり得る。
  - 影響: 配列引数を持つ API を個別変換なしでテストダブルに記録できず、任意文字列を同一性に含む値オブジェクトではスタブ値やスパイ結果の照合が不安定になる。
  - 対応案: 配列引数は要素型・次元・境界を含めてキー化するか、サポート外として `Class UnitTestUtils` の明示エラーにする。`IEquatable` / オブジェクト系キーにも同じエスケープと型タグ規則を適用し、衝突テストを追加する。
  - 保留理由: 対応が困難であり、かつ、現状では必要になっていないため
  - 保留解除条件: 必要とされたら

## 対応しないと決定した事項

- [x] [spec] WorksheetService.WriteRange の配列サイズを対象範囲と照合する
  - 詳細: `WriteRange` は `ValuesArray` をそのまま `target_range.Value` に代入しており、配列の次元数、下限、行数、列数が `RangeBounds` と一致するかを確認していない。Excel 側の代入仕様により、不足要素が `#N/A` になったり、余剰要素が書き込まれない可能性がある。
  - 影響: 呼び出し側の配列整形ミスが明示エラーではなくシート上のデータ破壊として表面化する。表データ書き込みの基盤 API として、意図しない `#N/A` や欠落に気付きにくい。
  - 結論: `WriteRange` は Excel `Range.Value` 代入への薄いラッパーとして扱い、配列サイズや次元の厳密な事前検証は追加しない。細かな代入仕様は Excel に準じる現行契約を維持する。

- [x] [spec] WorksheetService の配列数式コピーでコピー元を破壊しない
  - 詳細: `WorksheetService.pCopyCellCore` は配列数式コピー時に `src_cell = formula_str` でコピー元セルを通常数式へ一時変更しているため、コピー先エラー時にコピー元が壊れる懸念があった。
  - 結論: 指摘誤り。失敗系ユニットテストでコピー先エラー時もコピー元の配列数式が維持されることを確認したため、破壊バグは存在しないものとして扱う。実装修正は行わない。

- [x] [ux] Lib_UnitTest の実行用一時モジュール削除を生成済みモジュールだけに限定する
  - 詳細: `pRemoveRuntimeRunnerModule` は `pRuntimeRunnerModuleName` と一致するモジュールだけでなく、`Tmp_UTRUN[A-Z]{22}` に一致する標準モジュールをすべて削除する。テスト実行前は `pRuntimeRunnerModuleName` が空のため、同じ命名規則に偶然一致した利用者作成モジュールや調査用モジュールも削除対象になる。
  - 影響: `UnitTestMain` を実行しただけで、テストランナーが作成したものではない VBA モジュールを VBProject から削除する可能性がある。テスト基盤がソースを破壊し得るため、共通モジュール開発時のリスクが高い。
  - 結論: 杞憂に近い確率であり、仕様として許容します。なお、テスト結果判定に関わらない事項についてはバグではなく `[ux]` として扱います。

- [x] [spec] WorksheetService.CopyRange のコピー先範囲を空範囲・複数セルのまま受け付けない
  - 詳細: `CopyRange` では `DestinationRangeBounds.IsEmpty` と `DestinationRangeBounds.IsCell` の検査がコメントアウトされているが、実処理では `DestinationRangeBounds.Row` / `Column` をコピー先左上として使い、`FinishRow` / `FinishColumn` は無視している。空範囲でも保持している開始行・列へコピーされ、複数セル範囲を渡しても形状不一致を検出しない。
  - 影響: 呼び出し側が「コピー先なし」または「この範囲へコピー」と考えて渡した `WorksheetRangeBounds` が、意図しない左上セル起点の上書きとして実行され得る。
  - 結論: `CopyRange` のコピー先は範囲ではなく左上セルアンカーとして扱う現行動作を仕様とする。空範囲や複数セル範囲を追加検査で拒否する対応は行わない。

- [x] [spec] WorksheetService.GetUsedRangeBounds で値なし書式セルを空扱いしない
  - 詳細: `pGetRawUsedRange` は UsedRange が 1 セルだけの場合、値が空で四辺の罫線がないことだけを見て空範囲としている。塗りつぶし、フォント、表示形式、コメント、ハイパーリンク、入力規則など、値以外の使用状態を確認していない。
  - 影響: 書式だけを持つセルを `GetUsedRangeBounds(GetRawRange:=True)` が空範囲として返し、`CopyRange(CopyNumberFormat:=True)` など UsedRange に依存する処理で書式のみの使用セルが落ちる可能性がある。
  - 結論: 罫線と値のないその他書式のみのセルは空扱いする現行動作を仕様とする。書式だけを使用範囲として保持する対応は行わない。
