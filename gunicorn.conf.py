# gunicorn.conf.py
# Workers handle HTTP only — all Claude API calls run in daemon threads,
# so the worker is never blocked waiting on the API. Timeout is a safety net.
workers = 1          # MUST be 1 — job store is in-memory, multiple workers = jobs lost
threads = 4          # handle up to 4 concurrent requests within the single worker
worker_class = "gthread"
timeout = 300          # 5 min — parse + verify can take 90-120s combined
graceful_timeout = 30
keepalive = 5
max_requests = 500
max_requests_jitter = 50
preload_app = False    # Must be False — JOBS dict must be shared within worker
