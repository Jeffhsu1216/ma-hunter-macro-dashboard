import http.server, socketserver, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
with socketserver.TCPServer(("", 4891), http.server.SimpleHTTPRequestHandler) as h:
    h.serve_forever()
