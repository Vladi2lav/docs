# Локальный сервер без кэша: браузер всегда берёт свежие файлы (иначе старые модули смешиваются с новой страницей).
import http.server

class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {**http.server.SimpleHTTPRequestHandler.extensions_map,
                      '.js': 'text/javascript', '.mjs': 'text/javascript', '.json': 'application/json', '.css': 'text/css'}

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()

if __name__ == '__main__':
    http.server.ThreadingHTTPServer(('', 8080), Handler).serve_forever()
