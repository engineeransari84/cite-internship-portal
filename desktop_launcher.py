"""
Desktop launcher for the CITE Internship Records app.

Runs the Streamlit server in-process (no separate `streamlit` CLI call needed,
which matters once this is frozen into a single .exe with PyInstaller) and
opens the default browser to it. Closing the console window / Ctrl+C stops
the server.
"""

import os
import socket
import sys
import threading
import time
import webbrowser


def resource_path(rel_path: str) -> str:
    """Resolve a path next to this script, or inside the PyInstaller bundle."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_path)


def free_port(preferred: int = 8765) -> int:
    for port in range(preferred, preferred + 50):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return preferred


def main():
    # Frozen builds pass CITE_DATA_DIR through the run.bat / shortcut that
    # launches this exe (see PACKAGING.md); left unset, app.py falls back to
    # a "data" folder next to the exe, or the last folder chosen in the UI.
    from streamlit import config as st_config
    from streamlit.web import bootstrap

    app_path = resource_path("app.py")
    port = free_port()

    st_config.set_option("server.headless", True)
    st_config.set_option("server.port", port)
    st_config.set_option("server.address", "localhost")
    st_config.set_option("browser.gatherUsageStats", False)
    st_config.set_option("global.developmentMode", False)

    def open_browser():
        time.sleep(1.5)
        webbrowser.open(f"http://localhost:{port}")

    threading.Thread(target=open_browser, daemon=True).start()

    print(f"CITE Internship Records starting on http://localhost:{port}")
    print("Close this window (or press Ctrl+C) to stop the app.")
    bootstrap.run(app_path, False, [], {})


if __name__ == "__main__":
    main()
