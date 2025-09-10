import os
import signal
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer
from socketserver import ThreadingMixIn


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True


class HelloWorldHandler(BaseHTTPRequestHandler):
    def log_message(self, message_format, *args):
        sys.stderr.write("%s - - [%s] %s\n" % (
            self.client_address[0],
            self.log_date_time_string(),
            message_format % args,
        ))

    def do_GET(self):
        if self.path == "/":
            self._write_text_response(200, "Hello, World!")
            return

        if self.path in ("/healthz", "/ready"):
            self._write_text_response(200, "ok")
            return

        self.send_error(404, "Not Found")

    def _write_text_response(self, status_code, body_text):
        encoded_body = body_text.encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded_body)))
        self.end_headers()
        self.wfile.write(encoded_body)


def run_server(bind_host, bind_port):
    server = ThreadingHTTPServer((bind_host, bind_port), HelloWorldHandler)

    def handle_sigterm(signum, frame):
        server.shutdown()

    signal.signal(signal.SIGTERM, handle_sigterm)
    print(f"Serving on http://{bind_host}:{bind_port}", flush=True)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    host = os.environ.get("HOST", "0.0.0.0")
    run_server(host, port)

