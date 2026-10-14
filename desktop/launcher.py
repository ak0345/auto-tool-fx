"""Desktop launcher: a small window that runs the app on this computer and opens it in the browser.

Double-clicking the built app starts this. It starts the web app as a second copy of itself
(with --serve) on a free local port, opens the browser when it's ready, and stops it again
when the window is closed. Nothing is reachable from other computers.
"""

import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path

FROZEN = getattr(sys, "frozen", False)
HERE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
APP = HERE / "app.py"
LOG = Path(tempfile.gettempdir()) / "automation_tool.log"
TITLE = "Automation Tool"

THEME = {"theme.base": "dark", "theme.primaryColor": "#8b5cf6", "theme.backgroundColor": "#07040d",
         "theme.secondaryBackgroundColor": "#161024", "theme.textColor": "#ece8f5"}


def load_opencv():
    """OpenCV's Python loader recurses inside a packaged app, so load its compiled module directly
    (the scanned-PDF reader needs it)."""
    import importlib.util
    found = sorted((HERE / "cv2").glob("cv2*.so")) + sorted((HERE / "cv2").glob("cv2*.pyd"))
    if not found:
        return
    spec = importlib.util.spec_from_file_location("cv2", found[0])
    module = importlib.util.module_from_spec(spec)
    sys.modules["cv2"] = module
    spec.loader.exec_module(module)


def parent_alive(pid):
    if os.name == "nt":
        import ctypes
        handle = ctypes.windll.kernel32.OpenProcess(0x00100000, False, pid)        # SYNCHRONIZE
        if not handle:
            return False
        try:
            return ctypes.windll.kernel32.WaitForSingleObject(handle, 0) != 0      # 0 = it has exited
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)
    try:
        os.kill(pid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True


def stop_with_window(pid):
    """However the window goes (Quit, Cmd+Q, a crash), the app stops with it."""
    def watch():
        while parent_alive(pid):
            time.sleep(2)
        os._exit(0)
    threading.Thread(target=watch, daemon=True).start()


def serve(port, parent=None):
    """Run the Streamlit app (this is the second copy, started by the window)."""
    if parent:
        stop_with_window(parent)
    if FROZEN:
        try:
            load_opencv()
        except Exception as e:                       # the app says scanned PDFs can't be read
            sys.modules.pop("cv2", None)
            print(f"OpenCV couldn't load: {e}", file=sys.stderr)
    sys.path.insert(0, str(HERE))
    import turnover, inputs, run_queue  # noqa: F401  (bundled with the app this way)
    from streamlit.web import cli
    options = {"server.port": port, "server.address": "127.0.0.1", "server.headless": "true",
               "server.maxUploadSize": 20, "server.fileWatcherType": "none", "global.developmentMode": "false",
               "browser.gatherUsageStats": "false", "client.toolbarMode": "minimal", **THEME}
    sys.argv = ["streamlit", "run", str(APP)] + [f"--{k}={v}" for k, v in options.items()]
    sys.exit(cli.main())


def free_port():
    for port in range(8501, 8600):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def is_up(port):
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", port)) == 0


def start_server(port):
    args = ["--serve", str(port), str(os.getpid())]
    cmd = [sys.executable, *args] if FROZEN else [sys.executable, __file__, *args]
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    log = open(LOG, "w", encoding="utf-8", errors="replace")
    return subprocess.Popen(cmd, stdout=log, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                            creationflags=flags)


def window():
    import tkinter as tk

    port = free_port()
    url = f"http://localhost:{port}"
    server = start_server(port)

    root = tk.Tk()
    root.title(TITLE)
    root.configure(bg="#07040d")
    root.resizable(False, False)
    tk.Label(root, text=TITLE, font=("Helvetica", 20, "bold"), fg="#d8b4fe", bg="#07040d").pack(padx=40, pady=(28, 6))
    status = tk.Label(root, text="Starting, this takes a few seconds...", font=("Helvetica", 12),
                      fg="#ece8f5", bg="#07040d", wraplength=340, justify="center")
    status.pack(padx=30, pady=(0, 18))
    buttons = tk.Frame(root, bg="#07040d")
    buttons.pack(pady=(0, 24))
    open_btn = tk.Button(buttons, text="Open in browser", width=16, state="disabled",
                         command=lambda: webbrowser.open(url))
    open_btn.pack(side="left", padx=6)

    def quit_app():
        if server.poll() is None:
            server.terminate()
            try:
                server.wait(5)
            except subprocess.TimeoutExpired:
                server.kill()
        root.destroy()

    tk.Button(buttons, text="Quit", width=10, command=quit_app).pack(side="left", padx=6)
    root.protocol("WM_DELETE_WINDOW", quit_app)
    if sys.platform == "darwin":
        root.createcommand("::tk::mac::Quit", quit_app)        # Cmd+Q

    def ready():
        status.config(text="Running. The tool is open in your browser.\n"
                           "Keep this window open while you use it; close it when you're done.")
        open_btn.config(state="normal")
        webbrowser.open(url)

    def failed():
        status.config(text=f"The tool couldn't start. Details are in:\n{LOG}", fg="#fca5a5")

    def wait_for_server():
        deadline = time.monotonic() + 180                # first start after install can be slow
        while time.monotonic() < deadline:
            if server.poll() is not None:
                return root.after(0, failed)
            if is_up(port):
                return root.after(0, ready)
            time.sleep(0.3)
        root.after(0, failed)

    threading.Thread(target=wait_for_server, daemon=True).start()
    root.mainloop()


def selftest():
    """Used by the build: exit 0 only if the app's code loads and the scanned-PDF reader starts."""
    if FROZEN:
        load_opencv()
    sys.path.insert(0, str(HERE))
    import turnover
    import streamlit  # noqa: F401
    ok = APP.exists() and turnover.ocr_available()
    print("selftest", "ok" if ok else "FAILED", flush=True)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    if sys.argv[1:2] == ["--selftest"]:
        selftest()
    if len(sys.argv) >= 3 and sys.argv[1] == "--serve":
        serve(int(sys.argv[2]), int(sys.argv[3]) if len(sys.argv) > 3 else None)
    else:
        window()
