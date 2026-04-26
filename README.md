# Mihiraki

シンプルで高品質なPDFビューアーアプリ。  
Windowsのフォントレンダリング（ClearType）に最適化した、クリアで読みやすいPDF閲覧を目指しています。

## 機能

- **タブ表示** — 複数のPDFを同時に開いてタブで切り替え
- **ドラッグ＆ドロップ** — ウィンドウにPDFをドロップして開く
- **コマンドライン対応** — 既定のアプリとして関連付け可能
- **スクロールモード** — 全ページを縦に並べてスクロール閲覧
- **見開き表示（Spread）** — 左綴じ・右綴じ（RTL）に対応
- **ズーム / フィット** — 拡大縮小・ページフィット・横幅フィット
- **目次パネル（TOC）** — PDFのしおり情報をサイドパネルに表示
- **PNG書き出し** — 全ページをPNGファイルに一括変換
- **AI要約（Ollama連携）** — ローカルLLMによるPDF要約
- **設定記憶** — 表示モード・ズーム設定を次回起動時に復元

もう少し詳しい情報
https://it-libero.com/aitools/3206

## 動作環境

- Windows 10 / 11
- Python 3.10 以上

## インストール

### 1. リポジトリをクローン

```bash
git clone https://github.com/AtsushiCafeK/Mihiraki-PDF-Viewer.git
cd mihiraki
```

### 2. 依存パッケージをインストール

pyenv + poetryで構築していたので、poetry installでも構築できます。
一応、pipでのインストール手順でも説明書いてます。

```bash
pip install PySide6 pypdfium2 pymupdf Pillow requests
```

### 3. 起動

```bash
python PDFiumViewer.py
```

PDFファイルをコマンドライン引数として渡すこともできます。

```bash
python PDFiumViewer.py document.pdf
```

## 使い方

### 基本操作

| 操作 | 方法 |
|---|---|
| PDFを開く | ツールバー「Open」/ ドラッグ＆ドロップ |
| ページ移動 | ◀ ▶ ボタン / スライダー / ←→ キー |
| 左右端クリック | 画面の左端・右端クリックでページ移動 |
| ズーム | ツールバー「Zoom +/-」または「Fit Page / Fit Width」 |
| タブを閉じる | タブの × ボタン |

### 表示モード

| ボタン | 説明 |
|---|---|
| **Scroll** | 全ページを縦に並べてスクロール閲覧するモード |
| **Spread** | 見開き表示（表紙は単ページ） |
| **RTL** | 右綴じ（見開きの左右反転） |
| **RTL Nav** | RTL時に進む/戻る方向も反転 |
| **TOC** | 目次パネルの表示/非表示 |

### PNG書き出し

ツールバーの「Export PNG」から保存先フォルダを選ぶと、全ページを `0001.png`, `0002.png` ... の形式で書き出します。

### AI要約（Ollama連携）

ツールバーの「Summarize」を使うと、[Ollama](https://ollama.com/) のローカルLLMでPDFを要約できます。

**事前準備:**

1. Ollama をインストールして起動
2. 使用するモデルをダウンロード（例: `gemma3:12b`）

```bash
ollama pull gemma3:12b
```

3. 「Summarize」ボタンからホスト・モデル名・ページ範囲を設定して実行

画像解析対応モデル（Visionモデル）を使うと、図・表を含めた要約精度が向上します。

## ライセンス

[MIT License](LICENSE.txt) © IT Libero
