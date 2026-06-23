#!/usr/bin/env python3
"""本地测试服务器，用于开发调试。启动后打开 http://localhost:8080"""
import http.server
import webbrowser
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))
PORT = 8080


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()


print(f"启动本地服务: http://localhost:{PORT}")
webbrowser.open(f"http://localhost:{PORT}")
http.server.test(HandlerClass=NoCacheHandler, port=PORT)
