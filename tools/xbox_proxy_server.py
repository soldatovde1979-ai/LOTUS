# xbox_proxy_server.py — HTTPS CONNECT-прокси с резолвом доменов через DoH xbox-dns.
# Запуск на a-rds-appp01:  py xbox_proxy_server.py [порт]   (по умолчанию 18080)
# Права администратора не требуются (порт выше 1024). Только stdlib.
import base64
import json
import socket
import struct
import sys
import threading
import time
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

DOH_URL = "https://xbox-dns.ru/dns-query"
CACHE_TTL_DEFAULT = 300

_cache = {}
_cache_lock = threading.Lock()


def resolve_via_doh(host):
    """IPv4-адрес для host через DoH xbox-dns (кэш с TTL). None — не удалось."""
    with _cache_lock:
        ip, exp = _cache.get(host, (None, 0))
        if ip and exp > time.time():
            return ip

    # RFC 8484: DNS-запрос (A) кодируется в base64url в параметре ?dns=
    q = struct.pack(">HHHHHH", 0x1234, 0x0100, 1, 0, 0, 0)
    for part in host.split("."):
        b = part.encode("ascii", "ignore")
        q += bytes([len(b)]) + b
    q += b"\x00" + struct.pack(">HH", 1, 1)
    url = DOH_URL + "?dns=" + base64.urlsafe_b64encode(q).rstrip(b"=").decode("ascii")
    try:
        req = urllib.request.Request(url, headers={"Accept": "application/dns-json"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        print(f"[doh] {host}: ошибка DoH-запроса: {e}")
        return None

    ip, ttl = None, CACHE_TTL_DEFAULT
    for ans in data.get("Answer") or []:
        if ans.get("type") == 1:
            ip = str(ans.get("data", "")).strip()
            ttl = int(ans.get("TTL", CACHE_TTL_DEFAULT))
            break
    if ip:
        with _cache_lock:
            _cache[host] = (ip, time.time() + max(10, min(ttl, 3600)))
        print(f"[doh] {host} -> {ip}")
    else:
        print(f"[doh] {host}: A-записи нет")
    return ip


def _pump(src, dst):
    try:
        while True:
            data = src.recv(65536)
            if not data:
                break
            dst.sendall(data)
    except OSError:
        pass
    finally:
        try:
            dst.shutdown(socket.SHUT_WR)
        except OSError:
            pass


def relay(client, remote):
    t1 = threading.Thread(target=_pump, args=(client, remote), daemon=True)
    t2 = threading.Thread(target=_pump, args=(remote, client), daemon=True)
    t1.start()
    t2.start()
    t1.join()
    t2.join()
    for s in (client, remote):
        try:
            s.close()
        except OSError:
            pass


class ProxyHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def do_CONNECT(self):
        host_port = self.path
        if ":" in host_port:
            host, _, port_s = host_port.rpartition(":")
            try:
                port = int(port_s)
            except ValueError:
                port = 443
        else:
            host, port = host_port, 443

        ip = resolve_via_doh(host)
        if not ip:
            self.send_error(502, "DNS failed")
            return
        try:
            remote = socket.create_connection((ip, port), timeout=10)
        except OSError as e:
            print(f"[proxy] {host}:{port} -> {ip}: подключение не удалось: {e}")
            self.send_error(502, "connect failed")
            return
        remote.settimeout(300)
        self.connection.settimeout(300)
        print(f"[proxy] CONNECT {host}:{port} -> {ip}")
        try:
            self.connection.sendall(b"HTTP/1.1 200 Connection Established\r\n\r\n")
        except OSError:
            remote.close()
            return
        relay(self.connection, remote)

    def log_message(self, fmt, *args):
        pass  # свои строки уже печатаем; служебный лог BaseHTTPRequestHandler не нужен


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 18080
    srv = ThreadingHTTPServer(("0.0.0.0", port), ProxyHandler)
    print(f"[proxy] слушаю 0.0.0.0:{port}, DoH: {DOH_URL}")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
