#!/usr/bin/env python3
"""
473SerialTerm — Simple program to send and receive serial

Quick test (no hardware): Tools → Open loopback (loop://), then Send "hello".
"""

from __future__ import annotations
import sys, os, time, threading
from dataclasses import dataclass
from datetime import datetime
from typing import Optional, List

import serial
import serial.tools.list_ports

import ctypes  # for Windows taskbar/pin icon consistency

# --- PyInstaller-friendly forced imports for pyserial URL handlers ---
import serial.urlhandler.protocol_loop    # noqa: F401
import serial.urlhandler.protocol_socket  # noqa: F401
import serial.urlhandler.protocol_rfc2217 # noqa: F401

def resource_path(rel_path: str) -> str:
    """Return absolute path to resource, works in dev and PyInstaller bundles."""
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, rel_path)

from PySide6.QtCore import Qt, QThread, Signal, Slot, QTimer, QSettings
from PySide6.QtGui import QAction, QTextCursor, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QSpinBox, QCheckBox, QGroupBox, QGridLayout,
    QLineEdit, QFileDialog, QPlainTextEdit, QStatusBar, QMessageBox, QProgressBar,
    QSplitter
)

# ---------------- Serial worker ----------------
@dataclass
class PortConfig:
    port: str = ""
    baudrate: int = 115200
    bytesize: int = 8        # 5,6,7,8
    parity: str = 'N'        # N,E,O,M,S
    stopbits: float = 1      # 1, 1.5, 2
    rtscts: bool = False
    xonxoff: bool = False
    dsrdtr: bool = False
    timeout_ms: int = 50

    def to_kwargs(self):
        parity_map = {'N': serial.PARITY_NONE, 'E': serial.PARITY_EVEN,
                      'O': serial.PARITY_ODD, 'M': serial.PARITY_MARK, 'S': serial.PARITY_SPACE}
        stop_map = {1: serial.STOPBITS_ONE, 1.5: serial.STOPBITS_ONE_POINT_FIVE, 2: serial.STOPBITS_TWO}
        byte_map = {5: serial.FIVEBITS, 6: serial.SIXBITS, 7: serial.SEVENBITS, 8: serial.EIGHTBITS}
        return dict(
            baudrate=self.baudrate,
            bytesize=byte_map.get(self.bytesize, serial.EIGHTBITS),
            parity=parity_map.get(self.parity, serial.PARITY_NONE),
            stopbits=stop_map.get(self.stopbits, serial.STOPBITS_ONE),
            rtscts=self.rtscts, xonxoff=self.xonxoff, dsrdtr=self.dsrdtr,
            timeout=self.timeout_ms/1000.0, write_timeout=2
        )

class SerialWorker(QThread):
    data_received = Signal(object)  # bytes
    status_changed = Signal(bool)
    error = Signal(str)
    modem_lines = Signal(bool, bool, bool, bool)  # CTS, DSR, RI, CD

    def __init__(self):
        super().__init__()
        self._cfg = PortConfig()
        self._ser: Optional[serial.SerialBase] = None
        self._running = False
        self._lock = threading.Lock()
        self._pending_reopen = False

    def apply_config_and_open(self, cfg: PortConfig):
        with self._lock:
            self._cfg = cfg
            self._pending_reopen = True

    @Slot(bytes)
    def write_bytes(self, payload: bytes):
        with self._lock:
            if self._ser and self._ser.is_open:
                try:
                    self._ser.write(payload); self._ser.flush()
                except Exception as e:
                    self.error.emit(f"Write failed: {e}")

    @Slot(bool)
    def set_rts(self, level: bool):
        with self._lock:
            if self._ser and self._ser.is_open:
                try: self._ser.rts = level
                except Exception as e: self.error.emit(f"RTS set failed: {e}")

    @Slot(bool)
    def set_dtr(self, level: bool):
        with self._lock:
            if self._ser and self._ser.is_open:
                try: self._ser.dtr = level
                except Exception as e: self.error.emit(f"DTR set failed: {e}")

    def close_port(self):
        with self._lock:
            if self._ser:
                try: self._ser.close()
                except Exception: pass
                self._ser = None
                self.status_changed.emit(False)

    def run(self):
        self._running = True
        last_modem_emit = 0.0
        while self._running:
            if self._pending_reopen:
                with self._lock:
                    self._pending_reopen = False
                    try:
                        if self._ser and self._ser.is_open: self._ser.close()
                        # Works for URLs (loop://) and normal device names.
                        self._ser = serial.serial_for_url(self._cfg.port, **self._cfg.to_kwargs())
                        self.status_changed.emit(True)
                    except Exception as e:
                        self._ser = None
                        self.status_changed.emit(False)
                        self.error.emit(f"Open failed: {e}")

            with self._lock:
                ser = self._ser

            if ser and ser.is_open:
                try:
                    data = ser.read(ser.in_waiting or 1)
                    if data:
                        self.data_received.emit(bytes(data))
                    now = time.time()
                    if now - last_modem_emit > 0.05:
                        # Some backends don't expose these; guard them
                        try:
                            cts = bool(getattr(ser, 'cts', False))
                            dsr = bool(getattr(ser, 'dsr', False))
                            ri  = bool(getattr(ser, 'ri',  False))
                            cd  = bool(getattr(ser, 'cd',  False))
                            self.modem_lines.emit(cts, dsr, ri, cd)
                        except Exception:
                            pass
                        last_modem_emit = now
                except Exception as e:
                    self.error.emit(f"I/O error: {e}")
                    time.sleep(0.05)
            else:
                time.sleep(0.05)
        self.close_port()

    def stop(self):
        self._running = False
        self.wait(1000)

# ---------------- UI helpers ----------------
PRINTABLE_SET = set(range(32, 127))

def bytes_to_hex_dump(data: bytes, base_offset: int = 0, width: int = 16) -> str:
    lines = []
    for i in range(0, len(data), width):
        chunk = data[i:i+width]
        hex_part = ' '.join(f"{b:02X}" for b in chunk)
        ascii_part = ''.join(chr(b) if b in PRINTABLE_SET else '.' for b in chunk)
        lines.append(f"{base_offset + i:08X}  {hex_part:<{width*3}}  |{ascii_part:>{width}}|")
    return '\n'.join(lines)

def apply_appenders(s: str, mode: str) -> str:
    return {'None': s, 'CR': s+'\r', 'LF': s+'\n', 'CRLF': s+'\r\n', 'NULL': s+'\x00'}.get(mode, s)

# ---------------- Tabs ----------------
class PortTab(QWidget):
    open_clicked = Signal(PortConfig)
    close_clicked = Signal()
    set_rts = Signal(bool)
    set_dtr = Signal(bool)

    def __init__(self):
        super().__init__(); self._build()

    def _build(self):
        layout = QVBoxLayout(self)
        box = QGroupBox("Port"); grid = QGridLayout(); box.setLayout(grid)

        self.port_combo = QComboBox(); self.port_combo.setEditable(True)  # allow loop://
        self.refresh_btn = QPushButton("Refresh"); self.refresh_btn.clicked.connect(self.refresh_ports)
        grid.addWidget(QLabel("Serial Port:"), 0, 0)
        grid.addWidget(self.port_combo, 0, 1, 1, 2)
        grid.addWidget(self.refresh_btn, 0, 3)

        self.baud_combo = QComboBox(); self.baud_combo.setEditable(True)
        self.baud_combo.addItems(["9600","19200","38400","57600","115200","230400","460800","921600"])
        self.baud_combo.setCurrentText("115200")
        self.bytesize_combo = QComboBox(); self.bytesize_combo.addItems(["5","6","7","8"]); self.bytesize_combo.setCurrentText("8")
        self.parity_combo = QComboBox(); self.parity_combo.addItems(["N","E","O","M","S"]); self.parity_combo.setCurrentText("N")
        self.stopbits_combo = QComboBox(); self.stopbits_combo.addItems(["1","1.5","2"]); self.stopbits_combo.setCurrentText("1")
        self.rtscts_chk = QCheckBox("RTS/CTS"); self.xonxoff_chk = QCheckBox("XON/XOFF"); self.dsrdtr_chk = QCheckBox("DSR/DTR")

        grid.addWidget(QLabel("Baud:"), 1, 0); grid.addWidget(self.baud_combo, 1, 1)
        grid.addWidget(QLabel("Data:"), 1, 2); grid.addWidget(self.bytesize_combo, 1, 3)
        grid.addWidget(QLabel("Parity:"), 2, 0); grid.addWidget(self.parity_combo, 2, 1)
        grid.addWidget(QLabel("Stop:"), 2, 2); grid.addWidget(self.stopbits_combo, 2, 3)
        grid.addWidget(self.rtscts_chk, 3, 0); grid.addWidget(self.xonxoff_chk, 3, 1); grid.addWidget(self.dsrdtr_chk, 3, 2)

        btn_row = QHBoxLayout()
        self.open_btn = QPushButton("Open"); self.close_btn = QPushButton("Close"); self.close_btn.setEnabled(False)
        btn_row.addWidget(self.open_btn); btn_row.addWidget(self.close_btn); btn_row.addStretch(1)

        modem = QGroupBox("Modem & Control Lines"); mgrid = QGridLayout(); modem.setLayout(mgrid)
        self.cts_lbl = QLabel("CTS: ?"); self.dsr_lbl = QLabel("DSR: ?"); self.ri_lbl = QLabel("RI: ?"); self.cd_lbl = QLabel("CD: ?")
        self.rts_chk = QCheckBox("Assert RTS"); self.dtr_chk = QCheckBox("Assert DTR")
        mgrid.addWidget(self.cts_lbl,0,0); mgrid.addWidget(self.dsr_lbl,0,1); mgrid.addWidget(self.ri_lbl,0,2); mgrid.addWidget(self.cd_lbl,0,3)
        mgrid.addWidget(self.rts_chk,1,0); mgrid.addWidget(self.dtr_chk,1,1)

        self.open_btn.clicked.connect(self._emit_open)
        self.close_btn.clicked.connect(self.close_clicked.emit)
        self.rts_chk.toggled.connect(self.set_rts.emit)
        self.dtr_chk.toggled.connect(self.set_dtr.emit)

        layout.addWidget(box); layout.addLayout(btn_row); layout.addWidget(modem); layout.addStretch(1)
        self.refresh_ports()

    def refresh_ports(self):
        self.port_combo.clear()
        for p in serial.tools.list_ports.comports():
            self.port_combo.addItem(f"{p.device} — {p.description}", p.device)

    def _emit_open(self):
        try:
            cfg = PortConfig(
                port=self.port_combo.currentData() or self.port_combo.currentText(),
                baudrate=int(self.baud_combo.currentText()),
                bytesize=int(self.bytesize_combo.currentText()),
                parity=self.parity_combo.currentText(),
                stopbits=float(self.stopbits_combo.currentText()),
                rtscts=self.rtscts_chk.isChecked(),
                xonxoff=self.xonxoff_chk.isChecked(),
                dsrdtr=self.dsrdtr_chk.isChecked(),
            )
            self.open_clicked.emit(cfg)
        except Exception as e:
            QMessageBox.critical(self, "Invalid settings", str(e))

    def set_open_state(self, is_open: bool):
        self.open_btn.setEnabled(not is_open)
        self.close_btn.setEnabled(is_open)

    def update_modem(self, cts: bool, dsr: bool, ri: bool, cd: bool):
        def fmt(n, v): return f"{n}: {'On' if v else 'Off'}"
        self.cts_lbl.setText(fmt('CTS', cts)); self.dsr_lbl.setText(fmt('DSR', dsr))
        self.ri_lbl.setText(fmt('RI', ri));   self.cd_lbl.setText(fmt('CD', cd))

class DisplayTab(QWidget):
    clear_clicked = Signal()
    save_clicked = Signal()
    pause_toggled = Signal(bool)

    def __init__(self):
        super().__init__()
        self.byte_count = 0
        self.paused = False

        # ASCII mode state
        self._cr_overwrite_mode = False     # set by '\r', cleared by newline
        self._line_has_ts = False           # has the current line been timestamped?
        self._line_ts_len = 0               # length of timestamp prefix (chars)

        # ASCII+HEX single-line builder state
        self._ax_active = False
        self._ax_hex_tokens = []            # ["48","65","6C",...]
        self._ax_ascii_chars = []           # ["H","e","l",".",...]
        self._ax_prefix = ""                # timestamp locked when AX line is created

        self._build()

    def _build(self):
        layout = QVBoxLayout(self)

        # Controls row
        top = QHBoxLayout()
        self.view_mode = QComboBox(); self.view_mode.addItems(["ASCII","HEX","ASCII+HEX"]); self.view_mode.setCurrentText("ASCII")
        self.timestamp_chk = QCheckBox("Timestamp")
        self.autoscroll_chk = QCheckBox("Autoscroll"); self.autoscroll_chk.setChecked(True)
        self.pause_btn = QPushButton("Pause")
        self.clear_btn = QPushButton("Clear")
        self.save_btn = QPushButton("Save Log…")
        top.addWidget(QLabel("View:")); top.addWidget(self.view_mode)
        top.addWidget(self.timestamp_chk); top.addWidget(self.autoscroll_chk)
        top.addStretch(1); top.addWidget(self.pause_btn); top.addWidget(self.clear_btn); top.addWidget(self.save_btn)

        # Terminal view
        self.out_text = QPlainTextEdit(); self.out_text.setReadOnly(True)
        self.out_text.setLineWrapMode(QPlainTextEdit.NoWrap)
        mono = self.out_text.font()
        try: mono.setFamilies(["Consolas","Menlo","Courier New","Monospace"])
        except Exception: pass
        self.out_text.setFont(mono)

        self.pause_btn.clicked.connect(self._toggle_pause)
        self.clear_btn.clicked.connect(self._emit_clear)
        self.save_btn.clicked.connect(self.save_clicked.emit)

        layout.addLayout(top)
        layout.addWidget(self.out_text)

    # ---------- helpers ----------
    def _doc_col(self) -> int:
        c = self.out_text.textCursor()
        return c.position() - c.block().position()

    def _at_line_start(self) -> bool:
        return self._doc_col() == 0

    def _timestamp(self) -> str:
        return f"[{datetime.now().strftime('%H:%M:%S.%f')[:-3]}] " if self.timestamp_chk.isChecked() else ""

    def _reset_line_ts(self):
        self._line_has_ts = False
        self._line_ts_len = 0

    def _maybe_stamp_before_visible(self):
        """
        Insert timestamp lazily: only if we're at true column 0 and this line
        has not been timestamped yet. If overwrite mode is active (after '\r'),
        the stamp lands at col 0 and the cursor is left after the stamp, so
        subsequent chars overwrite after the prefix.
        """
        if not self._line_has_ts and self._at_line_start():
            ts = self._timestamp()
            if ts:
                self.out_text.insertPlainText(ts)
                self._line_has_ts = True
                self._line_ts_len = len(ts)
                # Cursor is already after the inserted ts; that's where we want it.

    def _insert_overwrite_char(self, ch: str):
        """Replace one char at the cursor (append if at EOL)."""
        c = self.out_text.textCursor()
        block = c.block()
        col = c.position() - block.position()
        line_text = block.text()
        if 0 <= col < len(line_text):
            c.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1)
            c.insertText(ch)
        else:
            c.clearSelection()
            c.insertText(ch)
    
    def _hex_write_tokens(self, tokens: List[str]):
        """Write a run of hex tokens on the current line with spacing and a timestamp at col 0."""
        if not tokens:
            return
        # Timestamp lazily at col 0
        if self._at_line_start():
            ts = self._timestamp()
            if ts:
                self.out_text.insertPlainText(ts)
        else:
            # ensure a space between previous content and new tokens
            prev = self.out_text.textCursor().block().text()
            if prev and not prev.endswith(" "):
                self.out_text.insertPlainText(" ")
        self.out_text.insertPlainText(" ".join(tokens))

    # ----- ASCII+HEX single-line management -----
    def _start_ax_line_if_needed(self):
        if self._ax_active:
            return
        self.out_text.moveCursor(QTextCursor.End)
        if self.out_text.textCursor().block().text() or self._doc_col() != 0:
            self.out_text.insertPlainText("\n")

        self._ax_prefix = self._timestamp()
        if self._ax_prefix:
            self.out_text.insertPlainText(self._ax_prefix)

        self._ax_hex_tokens.clear()
        self._ax_ascii_chars.clear()
        self._ax_active = True

    def _update_ax_line(self):
        """Rewrite the entire current block as: <prefix><hex>  |  <ascii>."""
        self.out_text.moveCursor(QTextCursor.End)
        c = self.out_text.textCursor()
        hex_part = " ".join(self._ax_hex_tokens)
        ascii_part = "".join(self._ax_ascii_chars)
        line = f"{self._ax_prefix}{hex_part}  |  {ascii_part}" if (self._ax_prefix or hex_part or ascii_part) else ""
        c.beginEditBlock()
        c.movePosition(QTextCursor.StartOfBlock)
        c.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
        c.insertText(line)
        c.endEditBlock()

    # ---------- UI actions ----------
    def _toggle_pause(self):
        self.paused = not self.paused
        self.pause_btn.setText("Resume" if self.paused else "Pause")
        self.pause_toggled.emit(self.paused)

    def _emit_clear(self):
        self.byte_count = 0
        self.out_text.clear()
        self._cr_overwrite_mode = False
        self._reset_line_ts()
        # reset ASCII+HEX builder state
        self._ax_active = False
        self._ax_hex_tokens.clear()
        self._ax_ascii_chars.clear()
        self._ax_prefix = ""
        self.clear_clicked.emit()

    # ---------- main entry ----------
    @Slot(object)
    def append_data(self, data_obj):
        if self.paused:
            return
        data = bytes(data_obj)
        mode = self.view_mode.currentText()

        # If leaving ASCII+HEX, reset its state so next time starts clean
        if mode != "ASCII+HEX" and self._ax_active:
            self._ax_active = False
            self._ax_hex_tokens.clear()
            self._ax_ascii_chars.clear()
            self._ax_prefix = ""

        self.out_text.moveCursor(QTextCursor.End)

        if mode == "ASCII":
            i = 0
            while i < len(data):
                b = data[i]

                # ----- CR / LF handling -----
                if b == 0x0D:  # CR
                    if i + 1 < len(data) and data[i + 1] == 0x0A:
                        # CRLF -> newline at col 0, no immediate stamp
                        self.out_text.insertPlainText("\n")
                        self._cr_overwrite_mode = False
                        self._reset_line_ts()
                        i += 2
                        continue
                    # lone CR -> go to start of line, enter overwrite (no stamp yet)
                    c = self.out_text.textCursor()
                    c.movePosition(QTextCursor.StartOfBlock)
                    self.out_text.setTextCursor(c)
                    self._cr_overwrite_mode = True
                    # do not stamp; wait for first visible char
                    i += 1
                    continue

                if b == 0x0A:  # LF
                    if i + 1 < len(data) and data[i + 1] == 0x0D:
                        # LFCR -> newline at col 0, no immediate stamp
                        self.out_text.insertPlainText("\n")
                        self._cr_overwrite_mode = False
                        self._reset_line_ts()
                        i += 2
                        continue
                    # lone LF: newline AND keep current column (no stamp)
                    keep_col = self._doc_col()
                    self.out_text.insertPlainText("\n")
                    self._cr_overwrite_mode = False
                    self._reset_line_ts()
                    if keep_col > 0:
                        self.out_text.insertPlainText(" " * keep_col)
                    i += 1
                    continue

                # ----- normal chars / tab / other controls -----
                if b == 0x09:
                    ch = "\t"
                elif 32 <= b <= 126:
                    ch = chr(b)
                else:
                    ch = f"\\x{b:02X}"

                # Lazy stamp happens right before we actually render a visible char
                self._maybe_stamp_before_visible()

                if self._cr_overwrite_mode:
                    self._insert_overwrite_char(ch)
                else:
                    self.out_text.insertPlainText(ch)

                i += 1

        elif mode == "HEX":
            # Print hex bytes; on LF or CRLF, also break the line so timestamps refresh.
            i = 0
            tokens: List[str] = []
            while i < len(data):
                b = data[i]
                # CRLF -> print "0D 0A" then newline (single break)
                if b == 0x0D and i + 1 < len(data) and data[i + 1] == 0x0A:
                    # flush current run plus the CRLF tokens
                    self._hex_write_tokens(tokens + ["0D", "0A"])
                    tokens.clear()
                    self.out_text.insertPlainText("\n")
                    i += 2
                    continue
                # lone LF -> print "0A" then newline
                if b == 0x0A:
                    self._hex_write_tokens(tokens + ["0A"])
                    tokens.clear()
                    self.out_text.insertPlainText("\n")
                    i += 1
                    continue
                # Normal byte
                tokens.append(f"{b:02X}")
                i += 1
            # flush any trailing tokens
            self._hex_write_tokens(tokens)

        else:  # ASCII+HEX
            # Build "<hex>  |  <ascii>" on a line; on LF or CRLF, finalize line and start a new one.
            i = 0
            while i < len(data):
                b = data[i]
                # Ensure we own a line with a timestamped prefix (once per line)
                self._start_ax_line_if_needed()

                if b == 0x0D and i + 1 < len(data) and data[i + 1] == 0x0A:
                    # Add CRLF tokens to the pair, update, then break line
                    self._ax_hex_tokens.extend(["0D", "0A"])
                    self._ax_ascii_chars.extend([".", "."])
                    self._update_ax_line()
                    self.out_text.insertPlainText("\n")
                    # reset builder so next byte starts a fresh pair line (fresh timestamp)
                    self._ax_active = False
                    self._ax_hex_tokens.clear()
                    self._ax_ascii_chars.clear()
                    self._ax_prefix = ""
                    i += 2
                    continue

                if b == 0x0A:
                    # Add LF token to the pair, update, then break line
                    self._ax_hex_tokens.append("0A")
                    self._ax_ascii_chars.append(".")
                    self._update_ax_line()
                    self.out_text.insertPlainText("\n")
                    self._ax_active = False
                    self._ax_hex_tokens.clear()
                    self._ax_ascii_chars.clear()
                    self._ax_prefix = ""
                    i += 1
                    continue

                # Normal byte extends the current pair line
                self._ax_hex_tokens.append(f"{b:02X}")
                self._ax_ascii_chars.append(chr(b) if 32 <= b <= 126 else ".")
                self._update_ax_line()
                i += 1


class SendTab(QWidget):
    send_text = Signal(bytes)

    def __init__(self):
        super().__init__()
        self._timer: Optional[QTimer] = None
        self._build()

    def _build(self):
        layout = QVBoxLayout(self)

        # Row 1: input + send
        row1 = QHBoxLayout()
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText("Type text to send… (escape sequences: \\r, \\n, \\x00, etc.)")
        self.send_btn = QPushButton("Send")
        self.input_edit.returnPressed.connect(self._send_clicked)
        row1.addWidget(self.input_edit, 1)
        row1.addWidget(self.send_btn)

        # Row 2: options (now includes Hex input toggle)
        row2 = QHBoxLayout()
        self.hex_mode_chk = QCheckBox("Hex input")  # <--- NEW
        self.append_combo = QComboBox(); self.append_combo.addItems(["None", "CR", "LF", "CRLF", "NULL"]); self.append_combo.setCurrentText("CRLF")
        self.encoding_combo = QComboBox(); self.encoding_combo.addItems(["utf-8", "latin-1", "cp1252"])
        self.repeat_chk = QCheckBox("Repeat every (ms):")
        self.repeat_spin = QSpinBox(); self.repeat_spin.setRange(10, 3600000); self.repeat_spin.setValue(0)
        self.repeat_chk.toggled.connect(self._toggle_repeat_enabled)
        self._toggle_repeat_enabled(False)

        row2.addWidget(self.hex_mode_chk)
        row2.addWidget(QLabel("Append:"))
        row2.addWidget(self.append_combo)
        row2.addWidget(QLabel("Encoding:"))
        row2.addWidget(self.encoding_combo)
        row2.addStretch(1)
        row2.addWidget(self.repeat_chk)
        row2.addWidget(self.repeat_spin)

        # Row 3: file send (unchanged)
        row3 = QHBoxLayout()
        self.file_path = QLineEdit(); self.file_path.setPlaceholderText("Select a file to send raw bytes…")
        self.browse_btn = QPushButton("Browse…")
        self.send_file_btn = QPushButton("Send File")
        self.chunk_spin = QSpinBox(); self.chunk_spin.setRange(1, 65536); self.chunk_spin.setValue(1024)
        self.inter_delay_spin = QSpinBox(); self.inter_delay_spin.setRange(0, 1000); self.inter_delay_spin.setValue(5)
        row3.addWidget(self.file_path, 1)
        row3.addWidget(self.browse_btn)
        row3.addWidget(QLabel("Chunk (B):"))
        row3.addWidget(self.chunk_spin)
        row3.addWidget(QLabel("Delay (ms):"))
        row3.addWidget(self.inter_delay_spin)
        row3.addWidget(self.send_file_btn)

        self.progress = QProgressBar(); self.progress.setValue(0); self.progress.setVisible(False)

        # Wire up
        self.send_btn.clicked.connect(self._send_clicked)
        self.browse_btn.clicked.connect(self._browse)
        self.send_file_btn.clicked.connect(self._send_file)
        self.hex_mode_chk.toggled.connect(self._on_hex_mode_toggled)  # <--- NEW

        layout.addLayout(row1)
        layout.addLayout(row2)
        layout.addLayout(row3)
        layout.addWidget(self.progress)
        layout.addStretch(1)

    def _on_hex_mode_toggled(self, on: bool):
        # In hex mode, we ignore encoding and change the placeholder hint
        self.encoding_combo.setEnabled(not on)
        if on:
            self.input_edit.setPlaceholderText("Type HEX bytes… e.g. 48 65 6C 6C 6F, 0x48,0x65, or 48656c6c6f")
        else:
            self.input_edit.setPlaceholderText("Type text to send… (escape sequences: \\r, \\n, \\x00, etc.)")

    def _toggle_repeat_enabled(self, enabled: bool):
        self.repeat_spin.setEnabled(enabled)

    def _send_clicked(self):
        try:
            if self.hex_mode_chk.isChecked():
                payload = self._parse_hex(self.input_edit.text())
                # Apply appenders as raw bytes
                payload += self._append_bytes()
            else:
                text = self.input_edit.text()
                # interpret Python-style escapes in TEXT mode
                try:
                    text = bytes(text, 'utf-8').decode('unicode_escape')
                except Exception:
                    pass
                text = self._apply_appenders_text(text)
                try:
                    payload = text.encode(self.encoding_combo.currentText(), errors='replace')
                except Exception:
                    payload = text.encode('utf-8', errors='replace')
        except ValueError as e:
            QMessageBox.critical(self, "Invalid hex input", str(e))
            return

        self.send_text.emit(payload)

        if self.repeat_chk.isChecked() and self.repeat_spin.value() > 0:
            if not self._timer:
                self._timer = QTimer(self); self._timer.timeout.connect(self._send_clicked)
            self._timer.start(self.repeat_spin.value())
        else:
            if self._timer:
                self._timer.stop()

    def _append_bytes(self) -> bytes:
        mode = self.append_combo.currentText()
        return {
            "None": b"",
            "CR": b"\r",
            "LF": b"\n",
            "CRLF": b"\r\n",
            "NULL": b"\x00",
        }.get(mode, b"")

    def _apply_appenders_text(self, s: str) -> str:
        mode = self.append_combo.currentText()
        if mode == "CR":   return s + "\r"
        if mode == "LF":   return s + "\n"
        if mode == "CRLF": return s + "\r\n"
        if mode == "NULL": return s + "\x00"
        return s

    def _parse_hex(self, s: str) -> bytes:
        """
        Parse flexible hex formats:
          - '48 65 6c 6c 6f'
          - '0x48,0x65,0x6c'
          - '48656c6c6f'
          - Mixed separators: spaces, commas, semicolons, newlines, underscores
          - Single nibble tokens (e.g., 'A') treated as '0A'
          - Even-length long tokens are split into pairs
        Raises ValueError on bad input or odd digit counts.
        """
        import re
        s = s.strip()
        if not s:
            return b""

        # Split on common separators; keep long tokens intact for pair-splitting
        tokens = re.split(r"[\s,;]+", s)
        out = bytearray()

        for tok in tokens:
            if not tok:
                continue
            # Normalize
            t = tok.lower().replace("0x", "").replace("_", "")
            # If it's purely hex?
            if not re.fullmatch(r"[0-9a-f]+", t):
                raise ValueError(f"Invalid hex token: '{tok}'")

            if len(t) == 1:
                # Single nibble -> 0x0X
                out.append(int("0" + t, 16))
            elif len(t) == 2:
                out.append(int(t, 16))
            else:
                # Longer runs like "deadbeef" -> split into pairs
                if len(t) % 2 != 0:
                    raise ValueError(f"Odd number of hex digits in token: '{tok}'")
                for i in range(0, len(t), 2):
                    out.append(int(t[i:i+2], 16))

        return bytes(out)

    def _browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select file to send")
        if path:
            self.file_path.setText(path)

    def _send_file(self):
        path = self.file_path.text().strip()
        if not path or not os.path.isfile(path):
            QMessageBox.warning(self, "No file", "Select a valid file first.")
            return
        try:
            total = os.path.getsize(path)
            sent = 0
            chunk = self.chunk_spin.value()
            delay = self.inter_delay_spin.value() / 1000.0
            self.progress.setVisible(True)
            self.progress.setRange(0, total)
            with open(path, 'rb') as f:
                while True:
                    buf = f.read(chunk)
                    if not buf:
                        break
                    self.send_text.emit(buf)
                    sent += len(buf)
                    self.progress.setValue(sent)
                    QApplication.processEvents()
                    if delay > 0:
                        time.sleep(delay)
            self.progress.setVisible(False)
        except Exception as e:
            self.progress.setVisible(False)
            QMessageBox.critical(self, "File send failed", str(e))

# ---------------- Main window ----------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("473SerialTerm"); self.resize(1000, 650)
        self.settings = QSettings("473SerialTerm", "473SerialTerm")

        self.worker = SerialWorker(); self.worker.start()

        self.display_tab = DisplayTab(); self.display_tab.setMinimumHeight(200)
        self.port_tab = PortTab(); self.send_tab = SendTab()
        self.tabs = QTabWidget(); self.tabs.addTab(self.port_tab, "Port"); self.tabs.addTab(self.send_tab, "Send")

        splitter = QSplitter(Qt.Vertical); splitter.addWidget(self.display_tab); splitter.addWidget(self.tabs)
        splitter.setCollapsible(0, False); splitter.setCollapsible(1, True)
        splitter.setStretchFactor(0, 3); splitter.setStretchFactor(1, 1); splitter.setSizes([600, 300])
        self.splitter = splitter; self.setCentralWidget(splitter)

        self.status = QStatusBar(); self.setStatusBar(self.status)
        # NEW: permanent connection state label
        self.conn_lbl = QLabel("CLOSED")
        self.conn_lbl.setStyleSheet("QLabel { color: #999; font-weight: 600; }")
        self.status.addPermanentWidget(self.conn_lbl)   # add first so it appears leftmost among permanents

        self.rx_count = 0; self.tx_count = 0
        self.rx_lbl = QLabel("RX: 0"); self.tx_lbl = QLabel("TX: 0")
        self.status.addPermanentWidget(self.rx_lbl)
        self.status.addPermanentWidget(self.tx_lbl)

        # NEW: remember last-open settings so we can show port @ baud
        self._last_cfg: Optional[PortConfig] = None

        self._build_menu()

        # Wiring
        self.port_tab.open_clicked.connect(self._open_port)
        self.port_tab.close_clicked.connect(self._close_port)
        self.port_tab.set_rts.connect(self.worker.set_rts)
        self.port_tab.set_dtr.connect(self.worker.set_dtr)
        self.send_tab.send_text.connect(self._tx_and_count)
        self.display_tab.clear_clicked.connect(self._clear_log)
        self.display_tab.save_clicked.connect(self._save_log)
        self.worker.data_received.connect(self._on_data)
        self.worker.status_changed.connect(self._on_status)
        self.worker.error.connect(self._on_error)
        self.worker.modem_lines.connect(self.port_tab.update_modem)

        self._raw_log = bytearray()
        self._restore_ui()

    def _build_menu(self):
        file_menu = self.menuBar().addMenu("&File")
        act_save = QAction("Save Display As…", self); act_save.triggered.connect(self._save_log); file_menu.addAction(act_save)
        file_menu.addSeparator(); act_quit = QAction("Quit", self); act_quit.triggered.connect(self.close); file_menu.addAction(act_quit)

        tools = self.menuBar().addMenu("&Tools")
        act_loop = QAction("Open loopback (loop://)", self); act_loop.triggered.connect(self._open_loopback); tools.addAction(act_loop)
        act_selftest = QAction("Self-test: generate incoming data", self); act_selftest.triggered.connect(self._self_test); tools.addAction(act_selftest)

        help_menu = self.menuBar().addMenu("&Help")
        act_about = QAction("About", self); act_about.triggered.connect(self._about); help_menu.addAction(act_about)

    def _about(self):
        QMessageBox.information(self, "About 473SerialTerm",
            "473SerialTerm — PySide6 + pyserial terminal.\n"
            "Tip: Tools → Open loopback to verify without hardware.")

    # Persistence
    def _restore_ui(self):
        last_port = self.settings.value("port.last", "")
        if last_port:
            idx = self.port_tab.port_combo.findData(last_port)
            if idx >= 0: self.port_tab.port_combo.setCurrentIndex(idx)
            else: self.port_tab.port_combo.setEditText(last_port)
        self.port_tab.baud_combo.setCurrentText(self.settings.value("port.baud", "115200"))
        self.port_tab.bytesize_combo.setCurrentText(self.settings.value("port.bytesize", "8"))
        self.port_tab.parity_combo.setCurrentText(self.settings.value("port.parity", "N"))
        self.port_tab.stopbits_combo.setCurrentText(self.settings.value("port.stopbits", "1"))
        self.port_tab.rtscts_chk.setChecked(self.settings.value("port.rtscts", False, bool))
        self.port_tab.xonxoff_chk.setChecked(self.settings.value("port.xonxoff", False, bool))
        self.port_tab.dsrdtr_chk.setChecked(self.settings.value("port.dsrdtr", False, bool))
        self.display_tab.view_mode.setCurrentText(self.settings.value("display.mode", "ASCII"))
        self.display_tab.timestamp_chk.setChecked(self.settings.value("display.ts", False, bool))
        self.send_tab.append_combo.setCurrentText(self.settings.value("send.append", "CRLF"))
        self.send_tab.encoding_combo.setCurrentText(self.settings.value("send.encoding", "utf-8"))
        sizes = self.settings.value("ui.splitterSizes")
        if sizes:
            try: self.splitter.setSizes([int(x) for x in sizes])
            except Exception: pass

    def _save_ui(self):
        self.settings.setValue("port.last", self.port_tab.port_combo.currentData() or self.port_tab.port_combo.currentText())
        self.settings.setValue("port.baud", self.port_tab.baud_combo.currentText())
        self.settings.setValue("port.bytesize", self.port_tab.bytesize_combo.currentText())
        self.settings.setValue("port.parity", self.port_tab.parity_combo.currentText())
        self.settings.setValue("port.stopbits", self.port_tab.stopbits_combo.currentText())
        self.settings.setValue("port.rtscts", self.port_tab.rtscts_chk.isChecked())
        self.settings.setValue("port.xonxoff", self.port_tab.xonxoff_chk.isChecked())
        self.settings.setValue("port.dsrdtr", self.port_tab.dsrdtr_chk.isChecked())
        self.settings.setValue("display.mode", self.display_tab.view_mode.currentText())
        self.settings.setValue("display.ts", self.display_tab.timestamp_chk.isChecked())
        self.settings.setValue("send.append", self.send_tab.append_combo.currentText())
        self.settings.setValue("send.encoding", self.send_tab.encoding_combo.currentText())
        self.settings.setValue("ui.splitterSizes", self.splitter.sizes())

    # Slots
    @Slot(PortConfig)
    def _open_port(self, cfg: PortConfig):
        self._save_ui()
        self._last_cfg = cfg            # NEW
        self.worker.apply_config_and_open(cfg)

    @Slot()
    def _close_port(self):
        self.worker.close_port()

    @Slot(bool)
    def _on_status(self, is_open: bool):
        self.port_tab.set_open_state(is_open)

        if is_open:
            # Prefer the last config we asked the worker to open
            port = getattr(self._last_cfg, "port", "?")
            baud = getattr(self._last_cfg, "baudrate", "?")
            self.conn_lbl.setText(f"OPEN — {port} @ {baud}  ({self._last_cfg.bytesize}{self._last_cfg.parity}{self._last_cfg.stopbits})")
            self.conn_lbl.setStyleSheet("QLabel { color: #0a0; font-weight: 600; }")
        else:
            self.conn_lbl.setText("CLOSED")
            self.conn_lbl.setStyleSheet("QLabel { color: #999; font-weight: 600; }")

    @Slot(object)
    def _on_data(self, data):
        b = bytes(data)
        self._raw_log.extend(b)
        self.rx_count += len(b)
        self.rx_lbl.setText(f"RX: {self.rx_count}")
        self.display_tab.append_data(b)

    @Slot(str)
    def _on_error(self, msg: str):
        self.status.showMessage(msg, 4000)

    def _tx_and_count(self, payload: bytes):
        self.worker.write_bytes(payload)
        self.tx_count += len(payload)
        self.tx_lbl.setText(f"TX: {self.tx_count}")

    def _clear_log(self):
        self._raw_log.clear(); self.rx_count = 0; self.tx_count = 0
        self.rx_lbl.setText("RX: 0"); self.tx_lbl.setText("TX: 0")

    def _save_log(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save display/log as…", filter="Text (*.txt);;Binary (*.bin);;All Files (*)")
        if not path: return
        try:
            if path.lower().endswith('.bin'):
                with open(path, 'wb') as f: f.write(self._raw_log)
            else:
                with open(path, 'w', encoding='utf-8', errors='replace') as f:
                    f.write(self.display_tab.out_text.toPlainText())
            self.status.showMessage(f"Saved to {path}", 3000)
        except Exception as e:
            QMessageBox.critical(self, "Save failed", str(e))

    # Tools
    def _open_loopback(self):
        cfg = PortConfig(
            port="loop://",
            baudrate=int(self.port_tab.baud_combo.currentText()),
            bytesize=int(self.port_tab.bytesize_combo.currentText()),
            parity=self.port_tab.parity_combo.currentText(),
            stopbits=float(self.port_tab.stopbits_combo.currentText()),
            rtscts=self.port_tab.rtscts_chk.isChecked(),
            xonxoff=self.port_tab.xonxoff_chk.isChecked(),
            dsrdtr=self.port_tab.dsrdtr_chk.isChecked(),
        )
        self._open_port(cfg)

    def _self_test(self):
        sample = b"Self-test: The quick brown fox jumps over the lazy dog.\r\n"
        for _ in range(5):
            self._on_data(sample); QApplication.processEvents(); time.sleep(0.02)

    def closeEvent(self, e):
        try: self._save_ui(); self.worker.stop()
        finally: super().closeEvent(e)

# --------------- main ---------------
def main():
    app = QApplication(sys.argv)

    # Windows taskbar grouping + correct pinned icon
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("com.473SerialTerm.app")
        except Exception:
            pass

    icon = QIcon(resource_path("assets/473SerialTermIcon.ico"))
    app.setWindowIcon(icon)   # default icon for all top-level windows

    win = MainWindow()
    win.setWindowIcon(icon)   # belt-and-suspenders: explicit on main window
    win.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
