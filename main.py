"""
VC Doc Importer — an embedded browser that imports a .docx into the VC.ru editor.

The whole app IS a browser window:
  * VC.ru loads right inside the window (QtWebEngine / real Chromium engine).
  * You log into your account on the page itself.
  * You open a new post -> the VC editor appears inside the same window.
  * Click into the article body, then press "Добавить документ" (or drag a
    .docx onto the window). The program then reproduces the article by sending
    REAL keyboard events into the editor, so VC's live markdown conversion
    fires exactly like manual typing:
        "## "  -> H2      "### " -> H3
        "- "   -> bullet  "1. "  -> numbered
        "> "   -> quote
        "**x**" -> bold   "*x*" -> italic
  * Images are placed on the clipboard and pasted (Paste action), so VC uploads
    them physically, in order, like you pasted them yourself.

Everything runs in one Qt event loop. Typing is driven by a QTimer that emits
one keystroke per tick, so the UI never freezes and you watch it fill live.
"""

import os
import sys
import datetime

from PySide6.QtCore import Qt, QUrl, QTimer, QEvent, QStandardPaths
from PySide6.QtGui import (
    QAction, QIcon, QKeyEvent, QImage, QGuiApplication, QPixmap
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QLineEdit, QPushButton, QFileDialog,
    QMessageBox, QDialog, QVBoxLayout, QHBoxLayout, QPlainTextEdit, QWidget,
    QLabel, QProgressBar
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEngineProfile, QWebEnginePage

from docx_parser import parse_docx, runs_to_markdown

APP_TITLE = "VC Doc Importer"
START_URL = "https://vc.ru/"


def asset(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "assets", name)


# JavaScript that focuses the article's contenteditable area and returns
# whether it was found. Sent before typing so key events land in the editor.
JS_FOCUS_EDITOR = r"""
(function () {
  var sels = [
    "div.l-entry__content [contenteditable='true']",
    "div.editor [contenteditable='true']",
    "[data-editor] [contenteditable='true']",
    "article [contenteditable='true']",
    "[contenteditable='true']"
  ];
  for (var i = 0; i < sels.length; i++) {
    var el = document.querySelector(sels[i]);
    if (el && el.offsetParent !== null) {
      el.focus();
      return true;
    }
  }
  return false;
})();
"""


class LogDialog(QDialog):
    """Copyable log window opened by the 'Лог' button."""

    def __init__(self, buffer, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Лог")
        self.resize(680, 440)
        lay = QVBoxLayout(self)

        row = QHBoxLayout()
        btn_copy = QPushButton("Копировать всё")
        btn_clear = QPushButton("Очистить")
        row.addWidget(btn_copy)
        row.addWidget(btn_clear)
        row.addStretch(1)
        lay.addLayout(row)

        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        self.text.setPlainText("\n".join(buffer))
        lay.addWidget(self.text)

        self._buffer = buffer
        btn_copy.clicked.connect(self._copy)
        btn_clear.clicked.connect(self._clear)

    def append(self, line):
        self.text.appendPlainText(line)

    def _copy(self):
        QGuiApplication.clipboard().setText(self.text.toPlainText())

    def _clear(self):
        self._buffer.clear()
        self.text.clear()


class WebPage(QWebEnginePage):
    """Custom page so we can accept drops and route JS console to the log."""

    def __init__(self, profile, log, parent=None):
        super().__init__(profile, parent)
        self._log = log

    def javaScriptConsoleMessage(self, level, message, line, source):
        # Keep noise low: only surface warnings/errors.
        if level != QWebEnginePage.JavaScriptConsoleMessageLevel.InfoMessageLevel:
            self._log(f"JS: {message}")


class TypingEngine:
    """
    Replays parsed actions into the focused editor by posting real key events
    to the QWebEngineView's focus proxy, one unit per timer tick. This keeps
    the UI responsive and lets you watch the article fill in live.
    """

    def __init__(self, view, log, on_progress, on_done):
        self.view = view
        self.log = log
        self.on_progress = on_progress
        self.on_done = on_done
        self.timer = QTimer()
        self.timer.setInterval(6)  # ms per keystroke — fast but visible
        self.timer.timeout.connect(self._tick)
        self._queue = []       # list of ("text", str) / ("key", Qt.Key) / ("image", path)
        self._total_blocks = 0
        self._done_blocks = 0
        self._busy_until_ms = 0

    # ---- build the keystroke queue from actions ----
    def load(self, actions):
        self._queue.clear()
        self._total_blocks = len(actions)
        self._done_blocks = 0
        for a in actions:
            t = a["type"]
            if t == "image":
                self._queue.append(("image", a["path"]))
                self._queue.append(("block", None))
                continue
            md = runs_to_markdown(a["runs"])
            if t == "heading":
                md = ("## " if a.get("level", 2) == 2 else "### ") + md
            elif t == "bullet":
                md = "- " + md
            elif t == "number":
                md = "1. " + md
            elif t == "quote":
                md = "> " + md
            # queue each character, then Enter, then a block marker
            for ch in md:
                self._queue.append(("text", ch))
            self._queue.append(("key", Qt.Key_Return))
            self._queue.append(("block", (t, md[:50])))

    def start(self):
        if not self._queue:
            self.on_done()
            return
        self.timer.start()

    def stop(self):
        self.timer.stop()

    def _target(self):
        # Key events must go to the internal render widget (focus proxy).
        return self.view.focusProxy()

    def _send_char(self, ch):
        w = self._target()
        if w is None:
            return
        key = 0
        press = QKeyEvent(QEvent.KeyPress, key, Qt.NoModifier, ch)
        release = QKeyEvent(QEvent.KeyRelease, key, Qt.NoModifier, ch)
        QApplication.postEvent(w, press)
        QApplication.postEvent(w, release)

    def _send_key(self, qtkey):
        w = self._target()
        if w is None:
            return
        press = QKeyEvent(QEvent.KeyPress, qtkey, Qt.NoModifier, "")
        release = QKeyEvent(QEvent.KeyRelease, qtkey, Qt.NoModifier, "")
        QApplication.postEvent(w, press)
        QApplication.postEvent(w, release)

    def _paste_image(self, path):
        """Put the image on the clipboard and issue a paste key combo."""
        img = QImage(path)
        if img.isNull():
            self.log(f"[!] Не удалось прочитать изображение: {os.path.basename(path)}")
            return
        QGuiApplication.clipboard().setImage(img)
        w = self._target()
        if w is None:
            return
        # Ctrl+V
        press = QKeyEvent(QEvent.KeyPress, Qt.Key_V, Qt.ControlModifier, "")
        release = QKeyEvent(QEvent.KeyRelease, Qt.Key_V, Qt.ControlModifier, "")
        QApplication.postEvent(w, press)
        QApplication.postEvent(w, release)
        self.log(f"Вставляю изображение: {os.path.basename(path)}")

    def _tick(self):
        import time
        now = int(time.time() * 1000)
        if now < self._busy_until_ms:
            return
        if not self._queue:
            self.timer.stop()
            self.log("Готово. Проверьте статью и опубликуйте.")
            self.on_done()
            return
        kind, val = self._queue.pop(0)
        if kind == "text":
            self._send_char(val)
        elif kind == "key":
            self._send_key(val)
        elif kind == "image":
            self._paste_image(val)
            self._busy_until_ms = now + 1500  # give VC time to upload
        elif kind == "block":
            self._done_blocks += 1
            self.on_progress(self._done_blocks, self._total_blocks)
            if val:
                self.log(f"[{val[0]}] {val[1]}")


class Browser(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1200, 820)
        try:
            self.setWindowIcon(QIcon(asset("icon.png")))
        except Exception:
            pass

        self._log_buffer = []
        self._log_dialog = None

        # Persistent profile keeps your VC login between runs.
        data_dir = os.path.join(
            QStandardPaths.writableLocation(QStandardPaths.AppDataLocation),
            "VCDocImporter",
        )
        os.makedirs(data_dir, exist_ok=True)
        self.profile = QWebEngineProfile("vc_doc_importer", self)
        self.profile.setPersistentStoragePath(os.path.join(data_dir, "storage"))
        self.profile.setCachePath(os.path.join(data_dir, "cache"))
        self.profile.setPersistentCookiesPolicy(
            QWebEngineProfile.PersistentCookiesPolicy.ForcePersistentCookies
        )

        self.view = QWebEngineView(self)
        self.page = WebPage(self.profile, self.log, self.view)
        self.view.setPage(self.page)
        self.setCentralWidget(self.view)
        self.view.setUrl(QUrl(START_URL))
        self.view.urlChanged.connect(self._on_url_changed)
        self.view.loadFinished.connect(lambda ok: self.log(
            f"Загрузка: {'ok' if ok else 'ошибка'} — {self.view.url().toString()}"))

        self.engine = TypingEngine(self.view, self.log, self._on_progress, self._on_done)

        self._build_toolbar()
        self._build_statusbar()

        # Accept drag & drop of a .docx anywhere on the window.
        self.setAcceptDrops(True)

        self.log("Программа запущена. Войдите в аккаунт VC и откройте редактор поста.")

    # ---------- UI ----------
    def _build_toolbar(self):
        tb = QToolBar("nav")
        tb.setMovable(False)
        self.addToolBar(tb)

        back = QAction("←", self)
        back.triggered.connect(self.view.back)
        tb.addAction(back)

        fwd = QAction("→", self)
        fwd.triggered.connect(self.view.forward)
        tb.addAction(fwd)

        reload = QAction("⟳", self)
        reload.triggered.connect(self.view.reload)
        tb.addAction(reload)

        home = QAction("VC", self)
        home.triggered.connect(lambda: self.view.setUrl(QUrl(START_URL)))
        tb.addAction(home)

        self.url_bar = QLineEdit()
        self.url_bar.setPlaceholderText("Адрес…")
        self.url_bar.returnPressed.connect(self._go_url)
        tb.addWidget(self.url_bar)

        self.btn_add = QPushButton("Добавить документ")
        self.btn_add.clicked.connect(self.on_add)
        tb.addWidget(self.btn_add)

        self.btn_log = QPushButton("Лог")
        self.btn_log.clicked.connect(self.on_log)
        tb.addWidget(self.btn_log)

    def _build_statusbar(self):
        bar = self.statusBar()
        self.progress = QProgressBar()
        self.progress.setMaximumWidth(220)
        self.progress.setValue(0)
        self.status_label = QLabel("Готово к работе.")
        bar.addWidget(self.status_label, 1)
        bar.addPermanentWidget(self.progress)

    # ---------- navigation ----------
    def _go_url(self):
        u = self.url_bar.text().strip()
        if not u:
            return
        if not u.startswith("http"):
            u = "https://" + u
        self.view.setUrl(QUrl(u))

    def _on_url_changed(self, qurl):
        self.url_bar.setText(qurl.toString())

    # ---------- logging ----------
    def log(self, msg):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        self._log_buffer.append(line)
        if self._log_dialog is not None and self._log_dialog.isVisible():
            self._log_dialog.append(line)

    def on_log(self):
        if self._log_dialog is None:
            self._log_dialog = LogDialog(self._log_buffer, self)
        self._log_dialog.show()
        self._log_dialog.raise_()
        self._log_dialog.activateWindow()

    # ---------- progress ----------
    def _on_progress(self, done, total):
        self.progress.setMaximum(total)
        self.progress.setValue(done)
        self.status_label.setText(f"Перенесено блоков: {done}/{total}")

    def _on_done(self):
        self.status_label.setText("Готово. Проверьте статью и опубликуйте.")
        self.btn_add.setEnabled(True)

    # ---------- import ----------
    def on_add(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите документ", "", "Word документ (*.docx)"
        )
        if path:
            self._import(path)

    def _import(self, path):
        if not path.lower().endswith(".docx"):
            QMessageBox.warning(self, APP_TITLE, "Нужен файл .docx")
            return
        if not os.path.exists(path):
            QMessageBox.critical(self, APP_TITLE, "Файл не найден.")
            return
        self.log(f"Читаю документ: {os.path.basename(path)}")
        try:
            actions, _ = parse_docx(path)
        except Exception as e:
            self.log(f"[!] Ошибка чтения документа: {e}")
            QMessageBox.critical(self, APP_TITLE, f"Не удалось прочитать документ:\n{e}")
            return
        n_txt = sum(1 for a in actions if a["type"] != "image")
        n_img = sum(1 for a in actions if a["type"] == "image")
        self.log(f"Разобрано: {n_txt} текстовых блоков, {n_img} изображений.")

        # Focus the editor first; only start typing if it was found.
        def after_focus(found):
            if not found:
                self.log("[!] Редактор не найден. Откройте создание поста и кликните "
                         "в поле статьи, затем повторите.")
                QMessageBox.information(
                    self, APP_TITLE,
                    "Не вижу поле статьи.\n\nОткройте новый пост на VC, кликните "
                    "в тело статьи (где мигает курсор), затем снова нажмите "
                    "«Добавить документ»."
                )
                return
            self.btn_add.setEnabled(False)
            self.status_label.setText("Переношу статью…")
            self.log("Начинаю перенос. Не трогайте страницу до конца.")
            self.engine.load(actions)
            self.engine.start()

        self.page.runJavaScript(JS_FOCUS_EDITOR, after_focus)

    # ---------- drag & drop ----------
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            for u in e.mimeData().urls():
                if u.toLocalFile().lower().endswith(".docx"):
                    e.acceptProposedAction()
                    return
        e.ignore()

    def dropEvent(self, e):
        for u in e.mimeData().urls():
            p = u.toLocalFile()
            if p.lower().endswith(".docx"):
                self._import(p)
                break


def main():
    # Needed so the persistent Chromium engine works well when frozen by PyInstaller.
    os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS", "--disable-gpu")
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    win = Browser()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
