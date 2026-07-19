export const SAMPLE_MARKDOWN = `---
marp: true
theme: default
paginate: true
title: Editable Marp
author: Marp PPTX
---

<!-- _class: lead -->

# Marp → Editable PPTX

**背景の忠実さ**と**コンテンツの編集性**を両立する

---

## 変換方針

- テーマ装飾は背景レイヤーとして保持
- タイトル・本文・リストはPowerPointテキストへ変換
- 通常画像は個別の編集可能要素として配置
- 未対応CSSは背景へ安全にフォールバック

---

![bg right:42%](/code-art.svg)

## Hybrid rendering

ブラウザが計算した位置・フォント・色を読み取り、同じ座標モデルからPPTXを生成します。

---

## Table support

| 項目 | 状態 | Editable |
| :--- | :--: | ---: |
| テキスト | 対応 | Yes |
| セル背景 | 対応 | Yes |
| 罫線 | 対応 | Yes |

---

## Inline image

通常配置した画像も、独立した画像要素として出力します。

![width:420px](/code-art.svg)

---

<!-- math: mathjax -->

## Math support

インライン数式 $E = mc^2$ とブロック数式を表示します。

$$
\\int_{-\\infty}^{\\infty} e^{-x^2}\\,dx = \\sqrt{\\pi}
$$
`

export const SAMPLE_CSS = `/* 任意のMarpテーマCSSをここに貼り付けられます */`
