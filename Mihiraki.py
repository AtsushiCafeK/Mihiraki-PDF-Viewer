# app.py (Mihiraki 1.0.9)
# Features added:
# 1) Open PDFs passed via command line (Open with / default app)
# 2) Drag & drop PDFs onto the window to open
# 3) Export PDF pages to PNG files (one file per page)
#
# Requirements:
#   pip install PySide6 pypdfium2 pymupdf Pillow requests

import sys
import os
import json
import base64
from io import BytesIO
from dataclasses import dataclass
from pathlib import Path
from collections import OrderedDict
from typing import Optional, Iterable

import requests
import pypdfium2 as pdfium
from PIL import Image
import fitz  # PyMuPDF

from PySide6.QtCore import (
    Qt, QThread, Signal, QCoreApplication, QPoint, QSettings
)
from PySide6.QtGui import (
    QAction, QImage, QPixmap, QPainter, QIcon, QShortcut, QKeySequence
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTabWidget, QWidget,
    QVBoxLayout, QToolBar, QSlider, QLabel, QHBoxLayout,
    QMessageBox, QDockWidget, QTreeWidget, QTreeWidgetItem,
    QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QToolButton,
    QDialog, QFormLayout, QLineEdit, QCheckBox, QSpinBox,
    QPushButton, QTextEdit, QProgressBar, QProgressDialog
)

# =============================
# App Identity
# =============================
APP_NAME = "Mihiraki"
APP_VERSION = "1.0.9"

WINDOWS_APP_ID = "jp.it-libero.mihiraki"
APP_ICON_REL = "assets/mihiraki.ico"

# =============================
# Viewer quality
# =============================
VIEW_QUALITY_OVERSAMPLE = 1.0
VIEW_CACHE_MAX_ITEMS = 64
SCROLL_PAGE_GAP = 8  # logical pixels between pages in scroll mode

# =============================
# Summary render quality
# =============================
SUMMARY_IMAGE_SCALE = 1.2
SUMMARY_JPEG_QUALITY = 75
SUMMARY_MAX_PAGES = 5000

# =============================
# Export PNG quality
# =============================
EXPORT_PNG_SCALE = 2.0  # 2.0〜3.0あたりが実用的。重いPDFなら2.0推奨

# =============================
# Ollama defaults
# =============================
DEFAULT_OLLAMA_HOST = "http://localhost:11434"
DEFAULT_MODEL = "gemma3:12b"
DEFAULT_BATCH_PAGES = 4

# =============================
# Settings keys
# =============================
SETTINGS_GROUP = "summarize"
KEY_HOST = "host"
KEY_MODEL = "model"
KEY_VISION = "vision"
KEY_BATCH = "batch_pages"

SETTINGS_GROUP_VIEW = "view"
KEY_VIEW_SCROLL = "scroll_mode"
KEY_VIEW_SPREAD = "spread"
KEY_VIEW_RTL = "rtl"
KEY_VIEW_RTL_NAV = "rtl_nav"
KEY_VIEW_ZOOM_MODE = "zoom_mode"
KEY_VIEW_ZOOM_FACTOR = "zoom_factor"
KEY_VIEW_TOC = "toc"


def set_appusermodel_id(app_id: str):
    """Windows taskbar identity."""
    if sys.platform != "win32":
        return
    try:
        import ctypes  # noqa
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def resource_path(rel: str) -> str:
    """PyInstaller resource path helper."""
    base = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(__file__).parent  # type: ignore[attr-defined]
    return str(base / rel)


def normalize_pdf_paths(paths: Iterable[str]) -> list[str]:
    out: list[str] = []
    for p in paths:
        try:
            pp = Path(p).expanduser()
            # macOSの「Open with」などは file:// を渡すケースがあるので一応対応
            s = str(pp)
            if s.lower().startswith("file://"):
                s = s[7:]
                pp = Path(s)
            if pp.exists() and pp.is_file() and pp.suffix.lower() == ".pdf":
                out.append(str(pp))
        except Exception:
            pass
    return out


class LRUCache:
    def __init__(self, max_items: int = 24):
        self.max_items = max_items
        self._d = OrderedDict()

    def get(self, key):
        if key not in self._d:
            return None
        self._d.move_to_end(key)
        return self._d[key]

    def put(self, key, value):
        self._d[key] = value
        self._d.move_to_end(key)
        while len(self._d) > self.max_items:
            self._d.popitem(last=False)

    def clear(self):
        self._d.clear()


@dataclass
class TocEntry:
    level: int
    title: str
    page_index: int  # 0-based


def qpixmap_from_pil(pil_img: Image.Image) -> QPixmap:
    img = pil_img.convert("RGB")
    w, h = img.size
    data = img.tobytes("raw", "RGB")
    qimg = QImage(data, w, h, 3 * w, QImage.Format.Format_RGB888).copy()
    return QPixmap.fromImage(qimg)


def jpeg_bytes_to_b64(jpeg_bytes: bytes) -> str:
    return base64.b64encode(jpeg_bytes).decode("ascii")


class OllamaClient:
    def __init__(self, host: str):
        self.host = host.rstrip("/")

    def chat(self, model: str, messages: list[dict], stream: bool = False, timeout_sec: int = 600) -> str:
        url = f"{self.host}/api/chat"
        payload = {"model": model, "messages": messages, "stream": stream}

        if not stream:
            r = requests.post(url, json=payload, timeout=timeout_sec)
            r.raise_for_status()
            data = r.json()
            return data.get("message", {}).get("content", "")

        r = requests.post(url, json=payload, stream=True, timeout=timeout_sec)
        r.raise_for_status()
        out = []
        for line in r.iter_lines():
            if not line:
                continue
            obj = json.loads(line.decode("utf-8"))
            if "message" in obj and "content" in obj["message"]:
                out.append(obj["message"]["content"])
            if obj.get("done"):
                break
        return "".join(out)


class SummarizeWorker(QThread):
    progress = Signal(int, str)
    finished_ok = Signal(str)
    finished_err = Signal(str)

    def __init__(
        self,
        pdf_path: str,
        start_page: int,
        end_page: int,
        ollama_host: str,
        model: str,
        use_vision: bool,
        batch_pages: int,
        parent=None
    ):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.start_page = start_page
        self.end_page = end_page
        self.ollama_host = ollama_host
        self.model = model
        self.use_vision = use_vision
        self.batch_pages = max(1, int(batch_pages))
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def _check_cancel(self):
        if self._cancel:
            raise RuntimeError("キャンセルされました。")

    def _extract_text(self, doc: fitz.Document, page_index: int) -> str:
        page = doc.load_page(page_index)
        return page.get_text("text").strip()

    def _render_page_jpeg_bytes(self, doc: pdfium.PdfDocument, page_index: int) -> bytes:
        page = doc.get_page(page_index)
        try:
            bitmap = page.render(scale=SUMMARY_IMAGE_SCALE, optimize_mode="lcd")
            pil_img = bitmap.to_pil().convert("RGB")
        finally:
            try:
                bitmap.close()
            except Exception:
                pass
            try:
                page.close()
            except Exception:
                pass

        buf = BytesIO()
        pil_img.save(buf, format="JPEG", quality=SUMMARY_JPEG_QUALITY, optimize=True)
        return buf.getvalue()

    def _make_map_prompt(self, page_infos: list[tuple[int, str]]) -> str:
        body_parts = []
        for idx, txt in page_infos:
            pno = idx + 1
            snippet = txt if txt else "(このページから抽出できるテキストが少ない/無し)"
            if len(snippet) > 12000:
                snippet = snippet[:12000] + "\n…(途中省略)…"
            body_parts.append(f"---\n[Page {pno}]\n{snippet}\n")
        body = "\n".join(body_parts)

        return (
            "あなたは資料読解の専門家です。以下はPDFの複数ページ分の内容です。\n"
            "目的: 『図表や画像の内容も含めて』重要点を落とさず、日本語で分かりやすく要約してください。\n\n"
            "出力形式:\n"
            "1) この範囲の要点（5〜12個の箇条書き）\n"
            "2) 図・表・画像が示す内容の説明（あれば）\n"
            "3) 重要な数値/条件/結論（あれば）\n"
            "4) 用語の補足（必要なら短く）\n\n"
            "内容:\n"
            f"{body}"
        )

    def _make_reduce_prompt(self, partial_summaries: list[str]) -> str:
        joined = "\n\n".join([f"---\n[Part {i+1}]\n{s}" for i, s in enumerate(partial_summaries)])
        return (
            "あなたは資料読解の専門家です。以下はPDFの部分要約の集合です。\n"
            "これらを統合して、重複を除き、全体として筋の良い最終要約を作ってください。\n"
            "必ず日本語で作ってください。\n\n"
            "必ず次の出力形式を守ってください。\n"
            "出力形式:\n"
            "A) 全体要約（10〜20行）\n"
            "B) 章/テーマ別まとめ（箇条書き）\n"
            "C) 重要な結論・意思決定ポイント\n"
            "D) ToDo / 次に確認すべき点（あれば）\n\n"
            "部分要約:\n"
            f"{joined}"
        )

    def _ollama_call(self, client: OllamaClient, prompt: str, images_b64: Optional[list[str]]) -> str:
        user_msg = {"role": "user", "content": prompt}
        if images_b64:
            user_msg["images"] = images_b64
        messages = [
            {"role": "system", "content": "あなたは正確で簡潔な要約を作るアシスタントです。根拠のない断定は避けてください。"},
            user_msg,
        ]
        return client.chat(model=self.model, messages=messages, stream=False, timeout_sec=900)

    def run(self):
        try:
            if self.end_page < self.start_page:
                raise RuntimeError("ページ範囲が不正です。")

            page_count_to_process = (self.end_page - self.start_page + 1)
            if page_count_to_process > SUMMARY_MAX_PAGES:
                raise RuntimeError(f"ページ数が多すぎます（{page_count_to_process}）。範囲を絞ってください。")

            client = OllamaClient(self.ollama_host)

            self.progress.emit(0, "PDFを開いています…")
            self._check_cancel()

            pdfium_doc = pdfium.PdfDocument(self.pdf_path)
            text_doc = fitz.open(self.pdf_path)

            partials: list[str] = []
            total_batches = (page_count_to_process + self.batch_pages - 1) // self.batch_pages

            for b in range(total_batches):
                self._check_cancel()
                batch_start = self.start_page + b * self.batch_pages
                batch_end = min(self.end_page, batch_start + self.batch_pages - 1)

                page_infos: list[tuple[int, str]] = []
                for p in range(batch_start, batch_end + 1):
                    self._check_cancel()
                    txt = self._extract_text(text_doc, p)
                    page_infos.append((p, txt))

                images_b64: Optional[list[str]] = None
                if self.use_vision:
                    images_b64 = []
                    for p in range(batch_start, batch_end + 1):
                        self._check_cancel()
                        jpeg_bytes = self._render_page_jpeg_bytes(pdfium_doc, p)
                        images_b64.append(jpeg_bytes_to_b64(jpeg_bytes))

                prompt = self._make_map_prompt(page_infos)
                pct = int((b / max(1, total_batches)) * 80)
                self.progress.emit(pct, f"部分要約中…（{b+1}/{total_batches}）")
                part = self._ollama_call(client, prompt, images_b64)
                partials.append(part.strip())

            self.progress.emit(85, "統合要約中…")
            reduce_prompt = self._make_reduce_prompt(partials)
            final_summary = self._ollama_call(client, reduce_prompt, images_b64=None)

            self.progress.emit(100, "完了")
            self.finished_ok.emit(final_summary.strip())
        except Exception as e:
            self.finished_err.emit(str(e))


class SummarizeDialog(QDialog):
    def __init__(self, page_count: int, current_page: int, settings: QSettings, parent=None):
        super().__init__(parent)
        self._settings = settings

        self.setWindowTitle(f"{APP_NAME} - Ollamaで要約")
        self.setModal(True)
        self.resize(610, 310)

        self._settings.beginGroup(SETTINGS_GROUP)
        saved_host = self._settings.value(KEY_HOST, DEFAULT_OLLAMA_HOST, type=str)
        saved_model = self._settings.value(KEY_MODEL, DEFAULT_MODEL, type=str)
        saved_vision = self._settings.value(KEY_VISION, True, type=bool)
        saved_batch = self._settings.value(KEY_BATCH, DEFAULT_BATCH_PAGES, type=int)
        self._settings.endGroup()

        self.host_edit = QLineEdit(saved_host, self)
        self.host_edit.setToolTip(
            "PCで起動したOllamaとMihirakiを接続するURLです。\n"
            "通常は http://localhost:11434 のままで動作します。"
        )

        self.model_edit = QLineEdit(saved_model, self)
        self.model_edit.setToolTip(
            "Ollamaでダウンロードしたモデル名を指定してください。\n"
            "画像も含めて要約する場合は、Vision（画像解析）対応モデル推奨です。"
        )

        self.vision_check = QCheckBox("画像も含めて要約（Visionモデル推奨）", self)
        self.vision_check.setChecked(bool(saved_vision))
        self.vision_check.setToolTip(
            "ページ画像も送って要約します。\n"
            "精度が上がる場合がありますが、処理が重くなることがあります。"
        )

        self.start_spin = QSpinBox(self)
        self.end_spin = QSpinBox(self)
        self.start_spin.setRange(1, max(1, page_count))
        self.end_spin.setRange(1, max(1, page_count))
        self.start_spin.setValue(1)
        self.end_spin.setValue(page_count)

        self.start_spin.setToolTip("要約する開始ページです。")
        self.end_spin.setToolTip("要約する終了ページです。")

        self.batch_spin = QSpinBox(self)
        self.batch_spin.setRange(1, 20)
        self.batch_spin.setValue(int(saved_batch) if int(saved_batch) > 0 else DEFAULT_BATCH_PAGES)
        self.batch_spin.setToolTip(
            "PDFを「数ページずつ」に分割して部分要約し、最後に統合します。\n"
            "この値は『一度に要約するページ数』です。\n"
            "小さいほど安定しやすい反面、回数が増えて時間がかかります。\n"
            "目安: 安定性優先 2〜4 / 速度優先 5〜8（モデルが強い時のみ）"
        )

        self.btn_current = QPushButton("現在ページだけ", self)
        self.btn_all = QPushButton("全ページ", self)
        self.btn_current.setToolTip("開始/終了を現在ページに設定します。")
        self.btn_all.setToolTip("開始/終了を全ページに設定します。")
        self.btn_current.clicked.connect(lambda: self._set_range_current(current_page))
        self.btn_all.clicked.connect(lambda: self._set_range_all(page_count))

        self.ok_btn = QPushButton("要約開始", self)
        self.cancel_btn = QPushButton("閉じる", self)
        self.ok_btn.setToolTip("設定した範囲で要約を開始します。")
        self.cancel_btn.setToolTip("要約せずに閉じます。")
        self.ok_btn.clicked.connect(self.accept)
        self.cancel_btn.clicked.connect(self.reject)

        form = QFormLayout()
        form.addRow("Ollama host", self.host_edit)
        form.addRow("Model", self.model_edit)
        form.addRow("", self.vision_check)

        range_row = QHBoxLayout()
        range_row.addWidget(QLabel("開始"))
        range_row.addWidget(self.start_spin)
        range_row.addSpacing(8)
        range_row.addWidget(QLabel("終了"))
        range_row.addWidget(self.end_spin)
        range_row.addSpacing(12)
        range_row.addWidget(self.btn_current)
        range_row.addWidget(self.btn_all)
        form.addRow("ページ範囲", range_row)

        form.addRow("バッチ（ページ/回）", self.batch_spin)

        btns = QHBoxLayout()
        btns.addStretch(1)
        btns.addWidget(self.ok_btn)
        btns.addWidget(self.cancel_btn)

        root = QVBoxLayout(self)
        root.addLayout(form)
        root.addStretch(1)
        root.addLayout(btns)

    def _set_range_current(self, current_page_0based: int):
        p = current_page_0based + 1
        self.start_spin.setValue(p)
        self.end_spin.setValue(p)

    def _set_range_all(self, page_count: int):
        self.start_spin.setValue(1)
        self.end_spin.setValue(page_count)

    def get_values(self):
        host = self.host_edit.text().strip() or DEFAULT_OLLAMA_HOST
        model = self.model_edit.text().strip() or DEFAULT_MODEL
        vision = self.vision_check.isChecked()
        batch_pages = int(self.batch_spin.value())

        self._settings.beginGroup(SETTINGS_GROUP)
        self._settings.setValue(KEY_HOST, host)
        self._settings.setValue(KEY_MODEL, model)
        self._settings.setValue(KEY_VISION, vision)
        self._settings.setValue(KEY_BATCH, batch_pages)
        self._settings.endGroup()

        return {
            "host": host,
            "model": model,
            "vision": vision,
            "start_page": self.start_spin.value() - 1,
            "end_page": self.end_spin.value() - 1,
            "batch_pages": batch_pages,
        }


class SummaryResultDialog(QDialog):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(860, 720)

        self.text = QTextEdit(self)
        self.text.setReadOnly(True)
        self.text.setToolTip("要約結果です。必要ならコピーしてメモなどに貼り付けてください。")

        self.progress = QProgressBar(self)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setToolTip("要約処理の進捗です。")

        self.status = QLabel("", self)

        self.btn_cancel = QPushButton("キャンセル", self)
        self.btn_close = QPushButton("閉じる", self)
        self.btn_close.setEnabled(False)

        self.btn_cancel.setToolTip("要約処理を中断します。")
        self.btn_close.setToolTip("ウィンドウを閉じます。")

        btns = QHBoxLayout()
        btns.addWidget(self.status, 1)
        btns.addWidget(self.btn_cancel, 0)
        btns.addWidget(self.btn_close, 0)

        root = QVBoxLayout(self)
        root.addWidget(self.text, 1)
        root.addWidget(self.progress, 0)
        root.addLayout(btns)

    def set_status(self, pct: int, message: str):
        self.progress.setValue(pct)
        self.status.setText(message)

    def set_result(self, content: str):
        self.text.setPlainText(content)
        self.btn_close.setEnabled(True)
        self.btn_cancel.setEnabled(False)

    def set_error(self, err: str):
        self.text.setPlainText("エラー:\n" + err)
        self.btn_close.setEnabled(True)
        self.btn_cancel.setEnabled(False)


class ClickNavGraphicsView(QGraphicsView):
    """
    - 左右端クリックで go_backward / go_forward
    - ドラッグ（パン）と両立（移動量が小さいときだけクリック判定）
    """
    def __init__(self, tab, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._tab = tab
        self._press_pos: Optional[QPoint] = None

        self.edge_px_min = 60
        self.edge_ratio = 0.08
        self.click_move_threshold_px = 4

        self.setToolTip(
            "PDF表示領域です。\n"
            "・ドラッグ: 表示を移動（パン）\n"
            "・左右端クリック: ページを戻る/進む\n"
            "・←→キー: ページを戻る/進む"
        )

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._press_pos = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self._press_pos is not None:
            release_pos = event.position().toPoint()
            moved = (release_pos - self._press_pos)

            if abs(moved.x()) <= self.click_move_threshold_px and abs(moved.y()) <= self.click_move_threshold_px:
                w = max(1, self.viewport().width())
                x = release_pos.x()
                edge = max(self.edge_px_min, int(w * self.edge_ratio))

                if x <= edge:
                    self._tab.go_backward()
                    event.accept()
                    self._press_pos = None
                    return
                if x >= (w - edge):
                    self._tab.go_forward()
                    event.accept()
                    self._press_pos = None
                    return

        self._press_pos = None
        super().mouseReleaseEvent(event)


class PdfTab(QWidget):
    """
    Spread:
      - 表紙（page 0）は単ページ
      - page 1 から見開きペア: (1,2), (3,4)...
    RTL:
      - 見開きの左右配置を反転（右綴じ/左綴じ表現）
    RTL Nav:
      - RTL時に操作方向も反転（紙の感覚に近づけるオプション）
    """
    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self.path = path

        self.pdfium_doc: Optional[pdfium.PdfDocument] = None
        self.toc_doc: Optional[fitz.Document] = None

        self.page_index = 0
        self.zoom_factor = 1.0
        self.zoom_mode = "custom"  # "custom" | "fit_page" | "fit_width"

        self.spread_enabled = False
        self.rtl_binding = False
        self.rtl_nav_reverse = False
        self.scroll_mode = False

        self._scroll_layout: list[tuple[float, float, float, float]] = []
        self._scroll_items: list[QGraphicsPixmapItem] = []
        self._scroll_scale: float = 1.0

        self.cache = LRUCache(max_items=VIEW_CACHE_MAX_ITEMS)

        self.scene = QGraphicsScene(self)
        self.view = ClickNavGraphicsView(self, self.scene, self)
        self.view.setDragMode(QGraphicsView.DragMode.ScrollHandDrag)
        self.view.setViewportUpdateMode(QGraphicsView.ViewportUpdateMode.FullViewportUpdate)
        self.view.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.SmoothPixmapTransform)

        self.btn_prev = QToolButton(self)
        self.btn_prev.setText("◀")
        self.btn_prev.setToolTip("戻る（見開き/ページ）")
        self.btn_prev.clicked.connect(self.go_backward)

        self.btn_next = QToolButton(self)
        self.btn_next.setText("▶")
        self.btn_next.setToolTip("進む（見開き/ページ）")
        self.btn_next.clicked.connect(self.go_forward)

        self.page_label = QLabel(self)
        self.page_label.setToolTip("現在のページ位置です。")

        self.slider = QSlider(Qt.Orientation.Horizontal, self)
        self.slider.valueChanged.connect(self._on_slider)
        self.slider.setToolTip("ドラッグでページを移動します（見開きON時も動作）。")

        bottom = QHBoxLayout()
        bottom.addWidget(self.btn_prev, 0)
        bottom.addWidget(self.btn_next, 0)
        bottom.addWidget(self.slider, 1)
        bottom.addWidget(self.page_label, 0)

        root = QVBoxLayout(self)
        root.setContentsMargins(6, 6, 6, 6)
        root.addWidget(self.view, 1)
        root.addLayout(bottom)

        self.toc_tree = QTreeWidget(self)
        self.toc_tree.setHeaderHidden(True)
        self.toc_tree.setToolTip("目次（PDFに目次がある場合に表示されます）。\nクリックでページ移動します。")
        self.toc_tree.itemClicked.connect(self._on_toc_clicked)

        self.view.verticalScrollBar().valueChanged.connect(self._on_scroll_bar_changed)

        self._load_documents(path)

    # meaning-based navigation
    def go_forward(self):
        if self.rtl_binding and self.rtl_nav_reverse:
            self.prev_page()
        else:
            self.next_page()

    def go_backward(self):
        if self.rtl_binding and self.rtl_nav_reverse:
            self.next_page()
        else:
            self.prev_page()

    def set_spread_enabled(self, enabled: bool):
        self.spread_enabled = bool(enabled)
        self.cache.clear()
        self.render_current_page()

    def set_rtl_binding(self, rtl: bool):
        self.rtl_binding = bool(rtl)
        self.cache.clear()
        self.render_current_page()

    def set_rtl_nav_reverse(self, enabled: bool):
        self.rtl_nav_reverse = bool(enabled)
        self._update_nav_buttons()

    def set_scroll_mode(self, enabled: bool):
        self.scroll_mode = bool(enabled)
        self.cache.clear()
        if self.scroll_mode:
            self._init_scroll_scene()
        else:
            self.render_current_page()

    def _init_scroll_scene(self):
        self.scene.clear()
        self._scroll_layout = []
        self._scroll_items = []

        n = self.page_count()
        if n <= 0:
            return

        eff_dpr = self._effective_dpr()
        vp_w = max(1, self.view.viewport().width())
        ref_pw, _ = self._page_size_points(0)
        self._scroll_scale = vp_w / max(1.0, ref_pw)

        y = 0.0
        max_w = 0.0
        for i in range(n):
            pw, ph = self._page_size_points(i)
            render_scale = self._scroll_scale * eff_dpr
            w_phys = max(1, round(pw * render_scale))
            h_phys = max(1, round(ph * render_scale))
            w_log = w_phys / eff_dpr
            h_log = h_phys / eff_dpr
            max_w = max(max_w, w_log)
            self._scroll_layout.append((0.0, y, w_log, h_log))

            ph_pix = QPixmap(w_phys, h_phys)
            ph_pix.setDevicePixelRatio(eff_dpr)
            ph_pix.fill(Qt.GlobalColor.lightGray)
            item = QGraphicsPixmapItem(ph_pix)
            item.setPos(0.0, y)
            self.scene.addItem(item)
            self._scroll_items.append(item)
            y += h_log + SCROLL_PAGE_GAP

        self.scene.setSceneRect(0.0, 0.0, max_w, y)
        self.view.resetTransform()

        self.view.verticalScrollBar().blockSignals(True)
        self._scroll_to_page(self.page_index)
        self.view.verticalScrollBar().blockSignals(False)

        self._render_visible_scroll_pages()
        self._update_page_label([self.page_index])
        self._update_nav_buttons()

    def _render_visible_scroll_pages(self):
        if not self.scroll_mode or not self._scroll_layout:
            return

        vr = self.view.mapToScene(self.view.viewport().rect()).boundingRect()
        buf = float(self.view.viewport().height())
        eff_dpr = self._effective_dpr()

        for i, (x, y, w, h) in enumerate(self._scroll_layout):
            if y + h < vr.top() - buf:
                continue
            if y > vr.bottom() + buf:
                break

            key = ('scroll', i, round(self._scroll_scale, 4))
            if self.cache.get(key) is not None:
                continue

            render_scale = self._scroll_scale * eff_dpr
            try:
                pil_img = self._render_page_pil(i, render_scale)
                qpix = qpixmap_from_pil(pil_img)
                qpix.setDevicePixelRatio(eff_dpr)
            except Exception:
                continue

            self.cache.put(key, qpix)
            if i < len(self._scroll_items):
                self._scroll_items[i].setPixmap(qpix)

    def _scroll_to_page(self, page_index: int):
        if page_index >= len(self._scroll_layout):
            return
        x, y, w, h = self._scroll_layout[page_index]
        self.view.ensureVisible(x, y, 1.0, 1.0, 0, 0)

    def _on_scroll_bar_changed(self, _value: int):
        if not self.scroll_mode:
            return
        self._render_visible_scroll_pages()
        self._update_scroll_page_index()

    def _update_scroll_page_index(self):
        if not self._scroll_layout:
            return
        vr = self.view.mapToScene(self.view.viewport().rect()).boundingRect()
        center_y = vr.center().y()

        best = 0
        best_dist = float('inf')
        for i, (x, y, w, h) in enumerate(self._scroll_layout):
            dist = abs((y + h * 0.5) - center_y)
            if dist < best_dist:
                best_dist = dist
                best = i

        if best != self.page_index:
            self.page_index = best
            self.slider.blockSignals(True)
            self.slider.setValue(best)
            self.slider.blockSignals(False)
            self._update_nav_buttons()

        self._update_page_label([best])

    def _spread_start(self, page_index: int) -> int:
        if page_index <= 0:
            return 0
        return page_index if ((page_index - 1) % 2 == 0) else (page_index - 1)

    def _display_pages(self, page_index: int) -> list[int]:
        n = self.page_count()
        if n <= 0:
            return []
        if not self.spread_enabled:
            return [max(0, min(n - 1, page_index))]
        if page_index == 0:
            return [0]
        start = self._spread_start(page_index)
        pages = [start]
        if start + 1 < n:
            pages.append(start + 1)
        return pages

    def _jump_prev_index(self) -> int:
        n = self.page_count()
        if n <= 0:
            return 0
        if not self.spread_enabled:
            return max(0, self.page_index - 1)
        if self.page_index <= 0:
            return 0
        if self.page_index == 1:
            return 0
        start = self._spread_start(self.page_index)
        prev_start = max(1, start - 2)
        return prev_start

    def _jump_next_index(self) -> int:
        n = self.page_count()
        if n <= 0:
            return 0
        if not self.spread_enabled:
            return min(n - 1, self.page_index + 1)
        if self.page_index == 0:
            return 1 if n > 1 else 0
        start = self._spread_start(self.page_index)
        next_start = start + 2
        if next_start >= n:
            return start
        return next_start

    def close_docs(self):
        try:
            if self.toc_doc is not None:
                self.toc_doc.close()
        except Exception:
            pass
        self.toc_doc = None
        self.pdfium_doc = None

    def page_count(self) -> int:
        return len(self.pdfium_doc) if self.pdfium_doc is not None else 0

    def _load_documents(self, path: str):
        try:
            self.pdfium_doc = pdfium.PdfDocument(path)
        except Exception as e:
            raise RuntimeError(f"PDFiumでPDFを開けません: {path}\n{e}")

        try:
            self.toc_doc = fitz.open(path)
        except Exception:
            self.toc_doc = None

        n = self.page_count()
        if n <= 0:
            raise RuntimeError(f"PDFページ数が0です: {path}")

        self.slider.blockSignals(True)
        self.slider.setMinimum(0)
        self.slider.setMaximum(n - 1)
        self.slider.setValue(0)
        self.slider.blockSignals(False)

        self.page_index = 0
        self._build_toc()
        self._update_nav_buttons()
        self.render_current_page()

    def _build_toc(self):
        self.toc_tree.clear()
        if self.toc_doc is None:
            return
        try:
            toc_raw = self.toc_doc.get_toc(simple=True)
        except Exception:
            toc_raw = []

        entries: list[TocEntry] = []
        for row in toc_raw:
            if len(row) >= 3:
                level, title, page_1based = row[0], row[1], row[2]
                if isinstance(page_1based, int) and page_1based >= 1:
                    entries.append(TocEntry(level=int(level), title=str(title), page_index=page_1based - 1))

        parents = {0: self.toc_tree.invisibleRootItem()}
        for ent in entries:
            parent_level = max(0, ent.level - 1)
            parent_item = parents.get(parent_level, self.toc_tree.invisibleRootItem())
            item = QTreeWidgetItem(parent_item, [ent.title])
            item.setData(0, Qt.ItemDataRole.UserRole, ent.page_index)
            parents[ent.level] = item

        self.toc_tree.expandToDepth(1)

    def _on_toc_clicked(self, item: QTreeWidgetItem, _col: int):
        page_index = item.data(0, Qt.ItemDataRole.UserRole)
        if isinstance(page_index, int):
            self.go_to_page(page_index)

    def _on_slider(self, page_index: int):
        self.go_to_page(page_index)

    def prev_page(self):
        self.go_to_page(self._jump_prev_index())

    def next_page(self):
        self.go_to_page(self._jump_next_index())

    def _update_nav_buttons(self):
        n = self.page_count()
        if n <= 0:
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(False)
            return

        if not self.spread_enabled:
            self.btn_prev.setEnabled(self.page_index > 0)
            self.btn_next.setEnabled(self.page_index < n - 1)
            return

        if self.page_index <= 0:
            self.btn_prev.setEnabled(False)
            self.btn_next.setEnabled(n > 1)
            return

        start = self._spread_start(self.page_index)
        self.btn_prev.setEnabled(True)
        self.btn_next.setEnabled((start + 2) < n)

    def go_to_page(self, page_index: int):
        n = self.page_count()
        if n <= 0:
            return
        page_index = max(0, min(n - 1, int(page_index)))
        self.page_index = page_index

        if self.slider.value() != page_index:
            self.slider.blockSignals(True)
            self.slider.setValue(page_index)
            self.slider.blockSignals(False)

        self._update_nav_buttons()

        if self.scroll_mode:
            self._scroll_to_page(page_index)
        else:
            self.render_current_page()

    def _update_page_label(self, displayed_pages: list[int]):
        if not displayed_pages:
            self.page_label.setText("")
            return
        if len(displayed_pages) == 1:
            self.page_label.setText(f"{displayed_pages[0] + 1} / {self.page_count()}")
        else:
            a, b = displayed_pages[0] + 1, displayed_pages[1] + 1
            self.page_label.setText(f"{a}-{b} / {self.page_count()}")

    def zoom_in(self):
        self.zoom_mode = "custom"
        self.zoom_factor = min(5.0, self.zoom_factor * 1.25)
        self.render_current_page()

    def zoom_out(self):
        self.zoom_mode = "custom"
        self.zoom_factor = max(0.25, self.zoom_factor * 0.8)
        self.render_current_page()

    def fit_page(self):
        self.zoom_mode = "fit_page"
        self.render_current_page()

    def fit_width(self):
        self.zoom_mode = "fit_width"
        self.render_current_page()

    def _effective_dpr(self) -> float:
        try:
            dpr = float(self.view.devicePixelRatioF())
        except Exception:
            dpr = 1.0
        return max(1.0, dpr) * float(VIEW_QUALITY_OVERSAMPLE)

    def _page_size_points(self, page_index: int) -> tuple[float, float]:
        assert self.pdfium_doc is not None
        page = self.pdfium_doc.get_page(page_index)
        try:
            w, h = page.get_size()
        finally:
            try:
                page.close()
            except Exception:
                pass
        return float(w), float(h)

    def _compute_fit_zoom_single(self, page_index: int) -> float:
        vp = self.view.viewport().size()
        vp_w = max(1, vp.width())
        vp_h = max(1, vp.height())

        page_w, page_h = self._page_size_points(page_index)
        page_w = max(1.0, page_w)
        page_h = max(1.0, page_h)

        if self.zoom_mode == "fit_width":
            return vp_w / page_w
        if self.zoom_mode == "fit_page":
            return min(vp_w / page_w, vp_h / page_h)
        return self.zoom_factor

    def _compute_fit_zoom_spread(self, left_idx: int, right_idx: int) -> float:
        vp = self.view.viewport().size()
        vp_w = max(1, vp.width())
        vp_h = max(1, vp.height())

        w1, h1 = self._page_size_points(left_idx)
        w2, h2 = self._page_size_points(right_idx)
        total_w = max(1.0, float(w1 + w2))
        max_h = max(1.0, float(max(h1, h2)))

        if self.zoom_mode == "fit_width":
            return vp_w / total_w
        if self.zoom_mode == "fit_page":
            return min(vp_w / total_w, vp_h / max_h)
        return self.zoom_factor

    def _render_page_pil(self, page_index: int, scale: float) -> Image.Image:
        assert self.pdfium_doc is not None
        page = self.pdfium_doc.get_page(page_index)
        try:
            bitmap = page.render(scale=scale, optimize_mode="lcd")
            pil_img = bitmap.to_pil().convert("RGB")
        finally:
            try:
                bitmap.close()
            except Exception:
                pass
            try:
                page.close()
            except Exception:
                pass
        return pil_img

    def _compose_spread(self, pil_a: Image.Image, pil_b: Image.Image, rtl: bool) -> Image.Image:
        if rtl:
            right = pil_a
            left = pil_b
        else:
            left = pil_a
            right = pil_b

        lw, lh = left.size
        rw, rh = right.size
        out_w = lw + rw
        out_h = max(lh, rh)

        canvas = Image.new("RGB", (out_w, out_h), (255, 255, 255))
        canvas.paste(left, (0, (out_h - lh) // 2))
        canvas.paste(right, (lw, (out_h - rh) // 2))
        return canvas

    def render_current_page(self):
        if self.pdfium_doc is None:
            return
        n = self.page_count()
        if n <= 0:
            return

        displayed = self._display_pages(self.page_index)

        eff_dpr = self._effective_dpr()
        dpr_key = round(eff_dpr, 3)

        if len(displayed) == 1:
            logical_zoom = self._compute_fit_zoom_single(displayed[0])
        else:
            logical_zoom = self._compute_fit_zoom_spread(displayed[0], displayed[1])

        logical_zoom = max(0.25, min(5.0, float(logical_zoom)))
        zoom_key = round(logical_zoom, 3)

        key = (
            tuple(displayed),
            zoom_key,
            self.zoom_mode,
            dpr_key,
            "spread" if (self.spread_enabled and len(displayed) == 2) else "single",
            "rtl" if self.rtl_binding else "ltr",
        )

        cached = self.cache.get(key)
        if cached is None:
            try:
                render_scale = logical_zoom * eff_dpr
                if len(displayed) == 1:
                    pil_img = self._render_page_pil(displayed[0], render_scale)
                else:
                    pil_a = self._render_page_pil(displayed[0], render_scale)
                    pil_b = self._render_page_pil(displayed[1], render_scale)
                    pil_img = self._compose_spread(pil_a, pil_b, rtl=self.rtl_binding)

                qpix = qpixmap_from_pil(pil_img)
                qpix.setDevicePixelRatio(eff_dpr)

            except Exception as e:
                QMessageBox.critical(self, "レンダリング失敗", f"PDFiumレンダリングに失敗:\n{e}")
                return

            self.cache.put(key, qpix)
            cached = qpix

        self.scene.clear()
        self.scene.addPixmap(cached)
        self.scene.setSceneRect(cached.rect())
        self.view.resetTransform()

        self._update_page_label(displayed)
        self._update_nav_buttons()

    def on_resize(self):
        if self.scroll_mode:
            self._init_scroll_scene()
        elif self.zoom_mode in ("fit_page", "fit_width"):
            self.render_current_page()


class MainWindow(QMainWindow):
    def __init__(self, app_icon: Optional[QIcon] = None, settings: Optional[QSettings] = None):
        super().__init__()
        self._settings = settings or QSettings()

        self._settings.beginGroup(SETTINGS_GROUP_VIEW)
        _saved_scroll = self._settings.value(KEY_VIEW_SCROLL, False, type=bool)
        _saved_spread = self._settings.value(KEY_VIEW_SPREAD, False, type=bool)
        _saved_rtl = self._settings.value(KEY_VIEW_RTL, False, type=bool)
        _saved_rtl_nav = self._settings.value(KEY_VIEW_RTL_NAV, False, type=bool)
        _saved_toc = self._settings.value(KEY_VIEW_TOC, True, type=bool)
        self._settings.endGroup()

        self.setWindowTitle(f"{APP_NAME} {APP_VERSION}")
        self.resize(1200, 800)

        if app_icon is not None:
            self.setWindowIcon(app_icon)

        # (2) Drag & Drop enable
        self.setAcceptDrops(True)

        self._status = self.statusBar()
        self._status.showMessage("Ready", 3000)

        self.tabs = QTabWidget(self)
        self.tabs.setDocumentMode(True)
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self._close_tab)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.setCentralWidget(self.tabs)

        self.toc_dock = QDockWidget("目次", self)
        self.toc_dock.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.toc_dock)

        tb = QToolBar("Main", self)
        tb.setMovable(False)
        self.addToolBar(tb)

        # --- Actions ---
        self.act_open = QAction("Open", self)
        self.act_open.triggered.connect(self.open_pdf_dialog)
        self._set_action_help(
            self.act_open,
            tooltip="PDFファイルを参照して開く",
            status="PDFファイルを参照して開きます（タブで複数開けます）。"
        )
        tb.addAction(self.act_open)

        # (3) Export PNG
        self.act_export_png = QAction("Export PNG", self)
        self.act_export_png.triggered.connect(self.export_png_all_pages)
        self._set_action_help(
            self.act_export_png,
            tooltip="PDFをページごとにPNGへ変換して保存する",
            status="保存フォルダを選び、全ページをPNGに変換してページ番号ファイル名で保存します。"
        )
        tb.addAction(self.act_export_png)

        tb.addSeparator()

        self.act_zoom_in = QAction("Zoom +", self)
        self.act_zoom_out = QAction("Zoom -", self)
        self.act_zoom_in.triggered.connect(self._zoom_in)
        self.act_zoom_out.triggered.connect(self._zoom_out)
        self._set_action_help(self.act_zoom_in, "拡大表示する", "表示を拡大します。")
        self._set_action_help(self.act_zoom_out, "縮小表示する", "表示を縮小します。")
        tb.addAction(self.act_zoom_in)
        tb.addAction(self.act_zoom_out)

        tb.addSeparator()

        self.act_fit_page = QAction("Fit Page", self)
        self.act_fit_width = QAction("Fit Width", self)
        self.act_fit_page.triggered.connect(self._fit_page)
        self.act_fit_width.triggered.connect(self._fit_width)
        self._set_action_help(
            self.act_fit_page,
            "PDFを常にウィンドウサイズへフィット表示する",
            "ページ全体が常にウィンドウ内に収まるように表示します。"
        )
        self._set_action_help(
            self.act_fit_width,
            "PDFを常にウィンドウサイズに合わせて横幅をフィット表示する",
            "横幅が常にウィンドウに合うように表示します（縦方向はスクロール）。"
        )
        tb.addAction(self.act_fit_page)
        tb.addAction(self.act_fit_width)

        tb.addSeparator()

        self.act_toc = QAction("TOC", self)
        self.act_toc.setCheckable(True)
        self.act_toc.setChecked(_saved_toc)
        self.toc_dock.setVisible(_saved_toc)
        self.act_toc.triggered.connect(self._toggle_toc)
        self._set_action_help(
            self.act_toc,
            "PDFに目次設定があれば表示する",
            "PDFに目次（しおり）がある場合、目次パネルを表示します。"
        )
        tb.addAction(self.act_toc)

        tb.addSeparator()

        self.act_spread = QAction("Spread", self)
        self.act_spread.setCheckable(True)
        self.act_spread.setChecked(_saved_spread)
        self.act_spread.triggered.connect(self._toggle_spread)
        self._set_action_help(
            self.act_spread,
            "PDFを見開き表示します。左綴じ、右ページが進む並び順にする",
            "PDFを見開き表示します（表紙は単ページ）。左綴じ・右ページが進む並び順です。"
        )
        tb.addAction(self.act_spread)

        self.act_rtl = QAction("RTL", self)
        self.act_rtl.setCheckable(True)
        self.act_rtl.setChecked(_saved_rtl)
        self.act_rtl.triggered.connect(self._toggle_rtl)
        self._set_action_help(
            self.act_rtl,
            "ページ綴じを右綴じ、左へページをめくる表示にする",
            "右綴じ（右→左へめくる）表示に切り替えます。見開きの左右配置が反転します。"
        )
        tb.addAction(self.act_rtl)

        self.act_rtl_nav = QAction("RTL Nav", self)
        self.act_rtl_nav.setCheckable(True)
        self.act_rtl_nav.setChecked(_saved_rtl_nav)
        self.act_rtl_nav.triggered.connect(self._toggle_rtl_nav)
        self._set_action_help(
            self.act_rtl_nav,
            "RTLに合わせた面付にページ並びを修正するオプション（RTL時は基本有効推奨）",
            "RTL表示に合わせて「進む／戻る」の操作方向も反転します。RTL時は基本有効推奨です。"
        )
        tb.addAction(self.act_rtl_nav)

        self.act_scroll = QAction("Scroll", self)
        self.act_scroll.setCheckable(True)
        self.act_scroll.setChecked(_saved_scroll)
        self.act_scroll.triggered.connect(self._toggle_scroll)
        self._set_action_help(
            self.act_scroll,
            "縦スクロールモードに切り替える",
            "スクロールモードをONにすると全ページが縦に並び、スクロールでページ移動できます。"
        )
        tb.addAction(self.act_scroll)

        tb.addSeparator()

        self.act_sum = QAction("Summarize", self)
        self.act_sum.triggered.connect(self.summarize_pdf)
        self._set_action_help(
            self.act_sum,
            "ローカルLLMによるAI要約機能（失敗時は再実行）",
            "ローカルLLM（Ollama）でAI要約します。モデルにより要約が揺れるため、失敗時は再実行してください。"
        )
        tb.addAction(self.act_sum)

        # Keyboard navigation (← →)
        self._install_keyboard_navigation()

        self._refresh_toc_panel()

        self._sum_worker: Optional[SummarizeWorker] = None
        self._sum_dialog: Optional[SummaryResultDialog] = None

    def _set_action_help(self, act: QAction, tooltip: str, status: str):
        act.setToolTip(tooltip)
        act.setStatusTip(status)
        act.hovered.connect(lambda s=status: self._status.showMessage(s, 6000))

    def _install_keyboard_navigation(self):
        sc_left = QShortcut(QKeySequence(Qt.Key_Left), self)
        sc_left.setAutoRepeat(True)
        sc_left.activated.connect(lambda: self._call_current(lambda t: t.go_backward()))

        sc_right = QShortcut(QKeySequence(Qt.Key_Right), self)
        sc_right.setAutoRepeat(True)
        sc_right.activated.connect(lambda: self._call_current(lambda t: t.go_forward()))

    # ===== (2) Drag & Drop =====
    def dragEnterEvent(self, event):
        md = event.mimeData()
        if md and md.hasUrls():
            # PDFが含まれるなら受け入れ
            for u in md.urls():
                if u.isLocalFile() and u.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        md = event.mimeData()
        if not md or not md.hasUrls():
            return
        paths = []
        for u in md.urls():
            if u.isLocalFile():
                p = u.toLocalFile()
                if p.lower().endswith(".pdf"):
                    paths.append(p)
        for p in normalize_pdf_paths(paths):
            self.open_pdf_path(p)
        event.acceptProposedAction()

    def current_tab(self) -> Optional[PdfTab]:
        w = self.tabs.currentWidget()
        return w if isinstance(w, PdfTab) else None

    def _call_current(self, fn):
        tab = self.current_tab()
        if tab is None:
            return
        fn(tab)

    def _refresh_toc_panel(self):
        tab = self.current_tab()
        if tab is None:
            self.toc_dock.setWidget(QWidget(self))
            self.toc_dock.setEnabled(False)
            return
        self.toc_dock.setWidget(tab.toc_tree)
        self.toc_dock.setEnabled(True)

    def _toggle_toc(self, checked: bool):
        self.toc_dock.setVisible(bool(checked))
        self._save_view_settings()

    def _toggle_spread(self, checked: bool):
        self._call_current(lambda t: t.set_spread_enabled(bool(checked)))
        self._save_view_settings()

    def _toggle_rtl(self, checked: bool):
        self._call_current(lambda t: t.set_rtl_binding(bool(checked)))
        self._save_view_settings()

    def _toggle_rtl_nav(self, checked: bool):
        self._call_current(lambda t: t.set_rtl_nav_reverse(bool(checked)))
        self._save_view_settings()

    def _toggle_scroll(self, checked: bool):
        self._call_current(lambda t: t.set_scroll_mode(bool(checked)))
        self._save_view_settings()

    def _save_view_settings(self):
        self._settings.beginGroup(SETTINGS_GROUP_VIEW)
        self._settings.setValue(KEY_VIEW_SCROLL, self.act_scroll.isChecked())
        self._settings.setValue(KEY_VIEW_SPREAD, self.act_spread.isChecked())
        self._settings.setValue(KEY_VIEW_RTL, self.act_rtl.isChecked())
        self._settings.setValue(KEY_VIEW_RTL_NAV, self.act_rtl_nav.isChecked())
        self._settings.setValue(KEY_VIEW_TOC, self.act_toc.isChecked())
        tab = self.current_tab()
        if tab is not None:
            self._settings.setValue(KEY_VIEW_ZOOM_MODE, tab.zoom_mode)
            self._settings.setValue(KEY_VIEW_ZOOM_FACTOR, tab.zoom_factor)
        self._settings.endGroup()
        self._settings.sync()

    def _zoom_in(self):
        self._call_current(lambda t: t.zoom_in())
        self._save_view_settings()

    def _zoom_out(self):
        self._call_current(lambda t: t.zoom_out())
        self._save_view_settings()

    def _fit_page(self):
        self._call_current(lambda t: t.fit_page())
        self._save_view_settings()

    def _fit_width(self):
        self._call_current(lambda t: t.fit_width())
        self._save_view_settings()

    # ===== (1) Open via args / open with =====
    def open_pdf_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf)")
        if not path:
            return
        self.open_pdf_path(path)

    def open_pdf_path(self, path: str):
        path = str(Path(path))
        if not (Path(path).exists() and Path(path).is_file()):
            QMessageBox.warning(self, "ファイルが見つかりません", f"存在しないファイルです:\n{path}")
            return
        if Path(path).suffix.lower() != ".pdf":
            QMessageBox.warning(self, "PDFではありません", f"PDFファイルではありません:\n{path}")
            return

        try:
            tab = PdfTab(path, parent=self)
        except Exception as e:
            QMessageBox.critical(self, "PDFを開けません", str(e))
            return

        title = Path(path).name
        self.tabs.addTab(tab, title)
        self.tabs.setCurrentWidget(tab)

        self._settings.beginGroup(SETTINGS_GROUP_VIEW)
        tab.zoom_mode = self._settings.value(KEY_VIEW_ZOOM_MODE, "fit_page", type=str)
        tab.zoom_factor = self._settings.value(KEY_VIEW_ZOOM_FACTOR, 1.0, type=float)
        self._settings.endGroup()

        tab.set_spread_enabled(self.act_spread.isChecked())
        tab.set_rtl_binding(self.act_rtl.isChecked())
        tab.set_rtl_nav_reverse(self.act_rtl_nav.isChecked())
        tab.set_scroll_mode(self.act_scroll.isChecked())

        self._refresh_toc_panel()
        self._status.showMessage("PDFを開きました。", 2500)

    def _close_tab(self, idx: int):
        w = self.tabs.widget(idx)
        self.tabs.removeTab(idx)
        if isinstance(w, PdfTab):
            w.close_docs()
        if w is not None:
            w.deleteLater()
        self._refresh_toc_panel()
        self._status.showMessage("タブを閉じました。", 2000)

    def _on_tab_changed(self, _idx: int):
        self._refresh_toc_panel()
        tab = self.current_tab()
        if tab is None:
            return

        self.act_spread.blockSignals(True)
        self.act_rtl.blockSignals(True)
        self.act_rtl_nav.blockSignals(True)
        self.act_scroll.blockSignals(True)

        self.act_spread.setChecked(tab.spread_enabled)
        self.act_rtl.setChecked(tab.rtl_binding)
        self.act_rtl_nav.setChecked(tab.rtl_nav_reverse)
        self.act_scroll.setChecked(tab.scroll_mode)

        self.act_spread.blockSignals(False)
        self.act_rtl.blockSignals(False)
        self.act_rtl_nav.blockSignals(False)
        self.act_scroll.blockSignals(False)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        tab = self.current_tab()
        if tab is not None:
            tab.on_resize()

    # ===== (3) Export PNG =====
    def export_png_all_pages(self):
        tab = self.current_tab()
        if tab is None:
            QMessageBox.information(self, "Export PNG", "まずPDFを開いてください。")
            return

        out_dir = QFileDialog.getExistingDirectory(self, "保存先フォルダを選択（ページごとにPNGを出力）", "")
        if not out_dir:
            return

        pdf_path = tab.path
        try:
            doc = pdfium.PdfDocument(pdf_path)
        except Exception as e:
            QMessageBox.critical(self, "Export PNG", f"PDFを開けません:\n{e}")
            return

        page_count = len(doc)
        if page_count <= 0:
            QMessageBox.warning(self, "Export PNG", "ページ数が0です。")
            return

        digits = max(4, len(str(page_count)))
        prog = QProgressDialog("PNGへ書き出し中…", "キャンセル", 0, page_count, self)
        prog.setWindowTitle("Export PNG")
        prog.setWindowModality(Qt.WindowModality.WindowModal)
        prog.setMinimumDuration(0)

        outp = Path(out_dir)
        try:
            outp.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Export PNG", f"保存先フォルダを作成できません:\n{e}")
            return

        try:
            for i in range(page_count):
                if prog.wasCanceled():
                    raise RuntimeError("キャンセルされました。")

                prog.setValue(i)
                prog.setLabelText(f"書き出し中… {i+1}/{page_count}")

                page = doc.get_page(i)
                try:
                    bitmap = page.render(scale=float(EXPORT_PNG_SCALE))
                    img = bitmap.to_pil().convert("RGB")
                finally:
                    try:
                        bitmap.close()
                    except Exception:
                        pass
                    try:
                        page.close()
                    except Exception:
                        pass

                filename = f"{(i+1):0{digits}d}.png"
                img.save(str(outp / filename), format="PNG", optimize=True)

                QApplication.processEvents()

            prog.setValue(page_count)
            QMessageBox.information(self, "Export PNG", f"完了しました。\n保存先: {out_dir}")
        except Exception as e:
            QMessageBox.warning(self, "Export PNG", str(e))

    def summarize_pdf(self):
        tab = self.current_tab()
        if tab is None:
            return

        n = tab.page_count()
        if n <= 0:
            return

        dlg = SummarizeDialog(page_count=n, current_page=tab.page_index, settings=self._settings, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        v = dlg.get_values()
        host = v["host"]
        model = v["model"]
        use_vision = bool(v["vision"])
        start_page = int(v["start_page"])
        end_page = int(v["end_page"])
        batch_pages = int(v["batch_pages"])

        out = SummaryResultDialog(title=f"{APP_NAME} - 要約結果", parent=self)
        out.set_status(0, "開始準備…")
        out.show()

        worker = SummarizeWorker(
            pdf_path=tab.path,
            start_page=start_page,
            end_page=end_page,
            ollama_host=host,
            model=model,
            use_vision=use_vision,
            batch_pages=batch_pages,
            parent=self
        )

        out.btn_cancel.clicked.connect(worker.cancel)
        out.btn_close.clicked.connect(out.accept)

        worker.progress.connect(lambda pct, msg: out.set_status(pct, msg))
        worker.finished_ok.connect(lambda text: out.set_result(text))
        worker.finished_err.connect(lambda err: out.set_error(err))

        self._sum_worker = worker
        self._sum_dialog = out
        worker.start()


def main():
    set_appusermodel_id(WINDOWS_APP_ID)

    QCoreApplication.setOrganizationName("IT-Libero")
    QCoreApplication.setApplicationName(APP_NAME)
    QCoreApplication.setApplicationVersion(APP_VERSION)

    app = QApplication(sys.argv)

    icon_path = resource_path(APP_ICON_REL)
    app_icon = QIcon(icon_path)
    app.setWindowIcon(app_icon)

    if hasattr(sys, "_MEIPASS"):
        _settings_dir = Path(sys.executable).parent  # exe化時はexeと同じフォルダ
    else:
        _settings_dir = Path(__file__).parent        # スクリプト実行時はスクリプトと同じフォルダ
    settings = QSettings(str(_settings_dir / "settings.ini"), QSettings.Format.IniFormat)

    w = MainWindow(app_icon=app_icon, settings=settings)
    w.show()

    # (1) Open PDFs passed via argv (Open with / default app)
    arg_paths = normalize_pdf_paths(sys.argv[1:])
    for p in arg_paths:
        w.open_pdf_path(p)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
