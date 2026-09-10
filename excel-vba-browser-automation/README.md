# Excel VBA ブラウザ入力自動化

Excel VBA + Windows API だけで、ブラウザを固定サイズで開き、実際にマウスカーソルを
動かしてクリックし、フィールドに数値や文字列を入力する自動化システムです。
外部ライブラリの参照設定は不要です（すべて `user32.dll` / `kernel32.dll` の標準API）。

## できること

- ブラウザ（Chrome、なければEdge）を指定サイズ・指定位置で起動し、リサイズ・最大化を禁止して固定
- `SetCursorPos` でマウスカーソルを実際に動かす（テレポートではなく滑らかに移動）
- 左クリック・ダブルクリック・右クリック・マウスホイールでのスクロールのシミュレーション
- `SendKeys` によるフィールドへの文字列・数値入力（既存値のクリアにも対応）
- ワークシート上に手順を並べるだけで動く簡易実行エンジン（VBA編集不要で手順を変更可能）
- 座標を調べるためのライブカーソル座標表示ツール

## 動作要件

- Windows + Excel 2010以降（VBA7 / PtrSafe 宣言を使用。32bit・64bit Office両対応）
- マクロを有効化（「トラストセンター」でマクロの実行を許可）
- Google Chrome または Microsoft Edge がインストール済み（レジストリ（App Paths）と
  インストール先候補パスの両方から自動検出。Chromeが無い場合はEdgeに自動フォールバック
  します。`DefaultBrowserPath()` をイミディエイトウィンドウで直接実行すると検出結果を
  確認できます。両方とも見つからない、または別ブラウザを使う場合は
  `LaunchChromeFixedSize` の `chromePath` 引数で明示的にパスを指定してください）

## ファイル構成

```
excel-vba-browser-automation/
  src/
    modWinAPI.bas            Windows API宣言・定数（共通）
    modBrowserWindow.bas     ブラウザ起動・ウィンドウサイズ固定
    modMouseKeyboard.bas     マウス移動/クリック・キー入力
    modAutomationRunner.bas  シート駆動の実行エンジン
    modCoordinateFinder.bas  座標確認用ツール
  SampleAutomationSteps.csv  手順シートのサンプルデータ
```

## セットアップ

1. Excelで新規ブック（マクロ有効ブック `.xlsm`）を作成
2. VBEditor（Alt+F11）を開き、`ファイル > ファイルのインポート` で
   `src/` 内の `.bas` を5つすべてインポート
   （`.bas` ファイル自体はコード中に日本語を含まないASCIIのみで書かれているため、
   文字コードの違いによる文字化けは起きません。インポート後に文字化けが見える場合は
   古いバージョンのファイルが残っている可能性があるので、一度モジュールを削除して
   最新の `.bas` を再インポートしてください）
3. ワークシートを1つ追加し、シート名を `AutomationSteps` に変更
4. 1行目に見出し（`Action, X, Y, Text, WaitMs, Url`）を入力し、
   `SampleAutomationSteps.csv` の内容を参考に手順を入力
   （CSVをそのまま `データ > テキストまたはCSVから` で取り込んでもよい）
5. VBEditorで `modAutomationRunner.RunAutomationSheet` を実行

## 手順シートの書き方

| 列 | 用途 |
|---|---|
| A: Action | 下表のアクション名 |
| B: X | クリック/移動先のX座標（画面絶対座標） |
| C: Y | クリック/移動先のY座標（画面絶対座標） |
| D: Text | TYPE / CLEARTYPE / KEY で入力する文字列 |
| E: WaitMs | WAIT の待機時間（ミリ秒） |
| F: Url | OPEN で開くURL |

対応アクション:

| Action | 内容 |
|---|---|
| `OPEN` | Chromeを起動し、B×C（幅×高さ）に固定してF列のURLを開く |
| `MOVE` | (X, Y) へカーソルを移動 |
| `CLICK` | (X, Y) を左クリック |
| `DBLCLICK` | (X, Y) をダブルクリック |
| `RIGHTCLICK` | (X, Y) を右クリック |
| `TYPE` | フォーカス中のフィールドにD列の文字列を入力 |
| `CLEARTYPE` | (X, Y) をクリック→既存値を全選択・削除→D列の文字列を入力 |
| `KEY` | D列の内容をそのまま `SendKeys` に渡す（例: `{ENTER}`, `{TAB}`, `^s`） |
| `SCROLL` | マウスホイールでスクロール。D列に「ノッチ数」（正=上, 負=下、例: `-3`）。B/C列を指定すると先にそこへカーソルを移動してからスクロール |
| `RESIZEACTIVE` | 現在フォーカスしているウィンドウをB×C（幅×高さ）にリサイズ。D列に `LOCK` と入れるとリサイズ・最大化も禁止して固定 |
| `WAIT` | E列のミリ秒だけ待機 |

## 選択中のウィンドウをリサイズする

`modBrowserWindow.ResizeForegroundWindow` を使うと、このプロジェクトで起動したブラウザ
に限らず、**今フォーカスしている任意のウィンドウ**を指定サイズにリサイズできます
（`GetForegroundWindow` APIで取得）。

```vb
' すぐにリサイズ（マクロ実行時にフォーカスしているウィンドウが対象）
Call ResizeForegroundWindow(1280, 800)

' 5秒後にリサイズ。VBEditorから実行する場合、実行直後はVBEditor自身が
' フォアグラウンドになってしまうため、この間に対象ウィンドウをクリックして
' フォーカスを移してください
Call ResizeForegroundWindow(1280, 800, delaySeconds:=5)

' リサイズ後、手動でのリサイズ・最大化も禁止して固定する
Call ResizeForegroundWindow(1280, 800, lockSize:=True, delaySeconds:=5)
```

手順シートからは `RESIZEACTIVE` アクションで同じことができます（上表を参照）。

## 座標の調べ方

`modCoordinateFinder.ShowCursorPositionLive` を実行してからブラウザ上の目的の
フィールド/ボタンにマウスを重ねると、Excelのステータスバーに現在のマウス座標が
リアルタイム表示されます。その値を手順シートのX/Yに使ってください。

```vb
Call ShowCursorPositionLive(30)  ' 30秒間、座標を表示し続ける（Escで中断）
```

## 注意事項・制約

- **座標は画面絶対座標**です。ディスプレイの拡大率（DPI）は100%に、Excelとブラウザの
  ウィンドウ配置は実行時と同じ状態にしておいてください。拡大率が異なると座標がずれます。
- 実行中はマウスを操作しないでください（`SetCursorPos` は実カーソルを動かすため、
  ユーザー操作と競合します）。
- 対象ページのレイアウトが変わると座標がずれて動作しなくなります。レイアウト変更が
  多いサイトを継続的に自動化するなら、要素をID/セレクタで指定できる
  **SeleniumBasic**（DOM要素ベース）の利用も検討してください。本実装はご要望の
  「カーソルを実際に動かしてクリックする」動作を優先した座標ベースの方式です。
- 自動化対象サイトの利用規約・robots.txt を確認し、許可されている範囲で使用してください。
  ログイン情報の平文管理やCAPTCHA回避目的での利用は避けてください。
