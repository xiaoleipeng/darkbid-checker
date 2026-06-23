#!/usr/bin/env python3
"""本地测试服务器，用于开发调试。启动后打开 http://localhost:8000"""
import http.server
import webbrowser
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))
PORT = 8080
print(f"启动本地服务: http://localhost:{PORT}")
webbrowser.open(f"http://localhost:{PORT}")
http.server.test(HandlerClass=http.server.SimpleHTTPRequestHandler, port=PORT)
