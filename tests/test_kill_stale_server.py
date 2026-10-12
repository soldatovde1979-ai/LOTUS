# -*- coding: utf-8 -*-
"""Тесты kill_stale_server(): остановка прошлых экземпляров и отказ при чужом порту."""
import os
import socket
import subprocess
import sys
import threading
import time
import unittest
from http.server import HTTPServer

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_gemini

PORT = 8767  # не боевой 8765, чтобы тест не зависел от работающего сервера


def spawn_server():
    """Запускает notes_gemini.py как отдельный процесс, возвращает Popen."""
    return subprocess.Popen(
        [sys.executable, "notes_gemini.py", str(PORT)],
        cwd=os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace")


def wait_listening(timeout=15):
    while timeout > 0:
        try:
            with socket.create_connection(("127.0.0.1", PORT), timeout=1):
                return True
        except OSError:
            time.sleep(0.3)
            timeout -= 0.3
    return False


class KillStaleServerTest(unittest.TestCase):
    def _stop(self, proc):
        proc.kill()
        proc.wait(timeout=10)

    def test_second_instance_kills_first_and_serves(self):
        """Второй запуск глушит первый и сам начинает работать."""
        first = spawn_server()
        try:
            self.assertTrue(wait_listening(), "первый экземпляр не поднялся")
            second = spawn_server()
            try:
                self.assertTrue(wait_listening(), "второй экземпляр не поднялся")
                self.assertIsNone(first.poll(), "первый экземпляр должен быть убит")
            finally:
                self._stop(second)
        finally:
            if first.poll() is None:
                self._stop(first)

    def test_outsider_on_port_is_not_killed(self):
        """Чужой процесс на порту не убивается, а запуск сервера прерывается."""
        blocker = HTTPServer(("127.0.0.1", PORT), notes_gemini.SimpleHTTPRequestHandler)
        t = threading.Thread(target=blocker.serve_forever, daemon=True)
        t.start()
        try:
            proc = spawn_server()
            try:
                out = proc.communicate(timeout=30)[0]
                self.assertNotEqual(proc.returncode, 0)
                self.assertIn("посторонним", out)
            except subprocess.TimeoutExpired:
                self.fail("сервер должен был прерваться, а не зависнуть")
            finally:
                if proc.poll() is None:
                    self._stop(proc)
        finally:
            blocker.shutdown()
            blocker.server_close()


if __name__ == "__main__":
    unittest.main()
