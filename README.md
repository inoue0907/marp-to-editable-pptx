# Marp to Editable PowerPoint

Marp Markdownをブラウザでレンダリングし、テーマ装飾を背景として保持しながら、主要コンテンツを編集可能なPowerPoint要素へ変換する独立Webアプリです。

## Features

- Marp CoreによるMarkdown・テーマCSSのレンダリング
- ブラウザとPowerPointで共通の1280×720座標モデル
- 見出し、本文、強調、箇条書きを編集可能なテキストへ変換
- 通常画像を独立した編集可能画像として配置
- Markdown表をPowerPointネイティブ表へ変換
- インライン・ブロック数式をPowerPointネイティブ数式（OMML）へ変換
- `![bg]`、疑似要素、テーマ装飾を背景レイヤーとして保持
- スピーカーノートに対応

## 起動

Windowsでは、Node.jsがPATHに登録されていない場合も含めて次のランチャーを使用できます。

```powershell
.\start-dev.ps1
```

実行ポリシーでブロックされた場合:

```powershell
powershell -ExecutionPolicy Bypass -File .\start-dev.ps1
```

または `start-dev.cmd` をダブルクリックしてください。

通常のNode.js環境では以下でも起動できます。

```bash
pnpm install
pnpm dev
```

npmを使用する場合:

```bash
npm install
npm run dev
```

## 現在の変換方針

- Marp Coreを唯一のMarkdown/CSSレンダラーとして使用
- ブラウザの確定座標とcomputed styleを中間モデルへ変換
- テーマ背景、疑似要素、背景画像、ページ番号は背景PNGとして保持
- 見出し、本文、リスト、通常画像、表、数式はPPTXネイティブ要素へ変換
- スピーカーノートをPPTXへ保存
- 変換できない要素は警告を表示し、背景レイヤーへ残す

## 制約

Webフォントとリモート画像は、配信元のCORS設定によって背景キャプチャできない場合があります。テーマで使用するフォントは、変換環境とPowerPoint閲覧環境へインストールしてください。PowerPointとブラウザのフォントメトリクス差により、改行位置がわずかに変わる場合があります。

## Native math

数式は画像化せず、KaTeXのMathMLをOMMLへ変換してPowerPointへ挿入します。出力後もPowerPointの数式として編集できます。
