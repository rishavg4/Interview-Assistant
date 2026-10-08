"""
start_server.py — Launches uvicorn keeping stdout/stderr flushed to a log file.
Run: python start_server.py
"""
import subprocess
import sys
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

proc = subprocess.Popen(
    [sys.executable, "-m", "uvicorn", "server:app",
     "--host", "127.0.0.1", "--port", "3001", "--log-level", "info"],
    stdout=sys.stdout,
    stderr=sys.stderr,
)
proc.wait()
