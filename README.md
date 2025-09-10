# Hello World (Python HTTP)

A tiny threaded HTTP server that returns "Hello, World!" at `/`.

## Run

```bash
python3 app.py
```

- Env vars:
  - `HOST` (default `0.0.0.0`)
  - `PORT` (default `8000`)

## Test

In a separate shell:

```bash
curl -i http://127.0.0.1:8000/
```

Expected response contains:

```
HTTP/1.0 200 OK
...
Hello, World!
```

Health endpoints:

```bash
curl -i http://127.0.0.1:8000/healthz
curl -i http://127.0.0.1:8000/ready
```