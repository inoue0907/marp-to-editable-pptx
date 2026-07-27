# Marpから編集可能なPowerPointへ

[English](./README.md)

Marp Markdownを、見た目を保ちながら編集可能なPowerPointへ変換します。

先にMarp自身へMarkdownとテーマCSSをレンダリングさせ、その完成DOMから最終的な座標とスタイルを取得します。複雑な装飾は選択できないスライド背景として保持し、主要コンテンツはPowerPointネイティブ要素として再構築します。

## 主な機能

- Marp標準テーマとカスタムテーマに対応
- Markdownで使用しているテーマを自動判定
- 見出し、本文、インライン装飾、上付き文字、下付き文字を編集可能なテキストとして出力
- 箇条書きと多階層番号リストをPowerPointネイティブのリストとして出力
- 最終的なフォントサイズ、色、行間、段落間隔、マーカーの色と大きさを反映
- 通常画像を独立したPowerPoint画像として出力
- ローカルのHTML5動画をPowerPointネイティブのメディアとして埋め込み
- Markdown表をPowerPointネイティブ表として出力
- インライン数式とブロック数式を編集可能なPowerPoint数式（OMML）として出力
- コードブロックを背景色付き・枠線なしの編集可能なテキストボックスとして出力
- `![bg]`、疑似要素、ロゴ、装飾図形、テーマ固有のアートワークをスライド背景へ保持
- Marpのスピーカーノートに対応

## 必要な環境

- Node.js LTS
- npmまたはpnpm
- Microsoft EdgeまたはGoogle Chrome
- Marpテーマが使用するフォント

変換自体にPowerPointは必要ありません。生成した`.pptx`の表示と編集にはPowerPointが必要です。

## インストール

```bash
git clone https://github.com/inoue0907/marp-to-editable-pptx.git
cd marp-to-editable-pptx
npm install
npm link
```

`npm link`により、短い`marp-pptx`コマンドがPATHへ登録されます。この作業は最初の1回だけ必要です。

インストールを確認します。

```bash
marp-pptx --help
```

## 使い方

以後は、このリポジトリ以外の任意のフォルダから実行できます。普段のMarpプロジェクトでターミナルを開き、Markdownを指定します。

```bash
marp-pptx slides.md
```

Markdownと同じフォルダへ同名のPowerPointを保存します。

```text
slides.md
slides.pptx
```

出力先を変更する場合：

```bash
marp-pptx slides.md --output dist/slides.pptx
```

通常はテーマを指定する必要はありません。明示的に上書きしたい場合だけ指定します。

```bash
marp-pptx slides.md --theme ./themes/my-theme.css
```

### CLIオプション

```text
marp-pptx <slides.md> [options]

-t, --theme <theme.css>    MarpテーマCSSを明示的に指定
-o, --output <file.pptx>   出力先
--html-output <file.html>  解決済みの自己完結Marp HTMLを保存
--html-only                PowerPointを作らずHTMLだけを保存
--callout-style <style>    引用をshape（既定）またはtextboxとして出力
--headed                   変換中のブラウザを表示
-h, --help                 ヘルプを表示
```

`--callout-style shape`では、引用の背景を編集可能な長方形、枠線を独立した
PowerPoint図形として出力します。`--callout-style textbox`では、編集可能な
テキストボックス自体に背景色を設定します。どちらも引用装飾をスライドの
背景画像へ埋め込みません。

## テーマの自動判定

`--theme`を省略すると、Markdown front matterの`theme:`を読み、次の順番でCSSを探します。

1. 最寄りの`.vscode/settings.json`にある`markdown.marp.themes`
2. Markdownと同じフォルダの`themes/<theme-name>.css`
3. コマンド実行フォルダの`themes/<theme-name>.css`
4. Markdownと同じフォルダの`theme.css`

VS Code設定については、Markdownのあるフォルダから上位へ`.vscode/settings.json`を探し、CSS内の`@theme`メタデータがfront matterと一致するテーマを使用します。

設定例：

```json
{
  "markdown.marp.themes": [
    "./themes/company.css",
    "./themes/technical.css"
  ]
}
```

```markdown
---
marp: true
theme: company
---
```

この状態なら、次のコマンドだけで`company.css`が自動的に選ばれます。

```bash
marp-pptx slides.md
```

Markdown内のローカル画像はMarkdownファイル基準、テーマCSS内の画像やフォントはCSSファイル基準で解決します。公開テーマでCSSの配置場所とアセットの相対パスが一致していない場合に備え、近い上位フォルダも探索します。

## VS Codeから使う

普段スライドを編集しているプロジェクトをVS Codeで開き、統合ターミナルから実行します。

```bash
marp-pptx slides.md
```

スライドのプロジェクトを、この変換ツールのリポジトリ内へ移す必要はありません。

現在開いているMarkdownを`Ctrl+Shift+B`で変換したい場合は、スライド側のプロジェクトへ`.vscode/tasks.json`を作成します。

```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "開いているMarpをPPTXへ変換",
      "type": "process",
      "command": "marp-pptx",
      "args": ["${file}"],
      "group": {
        "kind": "build",
        "isDefault": true
      },
      "problemMatcher": []
    }
  ]
}
```

設定後は、変換したいMarkdownを開いて`Ctrl+Shift+B`を押します。VS Codeが現在開いているファイルのパスを、PATHへ登録済みの`marp-pptx`へ渡します。

## Web画面を使う

開発用Web画面を起動します。

```bash
npm run dev
```

Windowsでは次のランチャーも利用できます。

```powershell
.\start-dev.ps1
```

実行ポリシーでブロックされた場合：

```powershell
powershell -ExecutionPolicy Bypass -File .\start-dev.ps1
```

または`start-dev.cmd`をダブルクリックします。

## 変換の仕組み

1. Marp CoreがMarkdownとテーマCSSをレンダリングします。
2. Marp Browserが1280×720の最終HTMLレイアウトを確定します。
3. 任意のCSSセレクタを独自解釈せず、完成DOMの座標とcomputed styleを読み取ります。
4. 文字、リスト、通常画像、表、コード、数式を一時的に除外し、残った装飾をキャプチャします。
5. 装飾を選択できないPowerPointスライド背景として埋め込みます。
6. 主要コンテンツを実測位置へPowerPointネイティブ要素として配置します。

テーマごとのCSS規則をPowerPoint側へ個別実装しないため、書き方の異なるカスタムテーマにも追従しやすい構成です。

## HTML出力

変換で使用する解決済みMarp HTMLを保存できます。

```bash
marp-pptx slides.md \
  --html-output ./dist/slides.html \
  --html-only
```

HTMLとPowerPointを両方作る場合は`--html-only`を外します。

## フォント

完成したMarp DOMのcomputed styleが選択したフォントをPowerPointでも使用します。フォントファイルをPPTXへ埋め込む処理は行いません。

必要なフォントは次の両方へインストールしてください。

- 変換を実行するPC
- フォント置換なしでPowerPointを開くすべてのPC

フォントがない場合はインストール済みフォントへフォールバックします。フォントが同じでも、ブラウザとPowerPointの文字メトリクス差により改行位置がわずかに変わる場合があります。

## 編集可能な要素と背景になる要素

PowerPointで編集可能：

- 見出しと本文
- 太字、斜体、文字色、ハイライト、上付き、下付き
- ネイティブの箇条書きと番号リスト
- 通常画像
- `<video src="...">`で指定したローカル動画
- 表
- インライン数式とブロック数式
- コードブロックの文字と背景色

スライド背景として保持：

- `![bg]`とMarpの高度背景
- ロゴとテーマアートワーク
- 疑似要素による装飾
- 装飾用の線、下線、影、複雑なCSS効果
- 編集対象ではないページ装飾

## 現在の制約

- ブラウザとPowerPointのフォントメトリクス差により、改行位置が完全には一致しない場合があります。
- フォント埋め込みには対応していません。
- ネットワークや配信元の制限により、リモートフォントやリモート画像を取得できない場合があります。
- 動画の再生可否は閲覧環境のコーデックに依存します。MP4/H.264を推奨します。
- 複雑なCSS効果は背景画像になります。
- コードのシンタックスハイライトは、ハイライター固有の効果をすべて再現できない場合があります。
- CSSの文字組み・レイアウト機能の一部にはPowerPoint側の対応機能がありません。

## 機密情報と外部通信

変換は一時的なローカルWebサーバーとインストール済みブラウザを使い、ローカルで実行します。このプロジェクトがファイルを外部サービスへアップロードすることはありません。

ただし、MarkdownまたはCSSに外部URLが書かれている場合、そのURLから画像やフォントを取得することがあります。機密スライドではローカルアセットを使用し、テーマCSSに外部URLがないか確認してください。

## CLIの更新

```bash
cd marp-to-editable-pptx
git pull
npm install
npm link
```

Node.jsを入れ替えた後などにコマンドが見つからなくなった場合は、もう一度`npm link`を実行してください。
